/* Access Control Management Page — Users, Roles, Permissions (Admin only) */

let acMatrix = {};
let acRoles = [];
let acPages = [];
let acRolesInfo = {};
let acUsers = [];

if (typeof window.showNotification !== "function") {
  window.showNotification = function(message, type) {
    try {
      const containerId = "acToastContainer";
      let container = document.getElementById(containerId);
      if (!container) {
        container = document.createElement("div");
        container.id = containerId;
        container.style.position = "fixed";
        container.style.top = "1rem";
        container.style.right = "1rem";
        container.style.zIndex = "1200";
        container.style.display = "flex";
        container.style.flexDirection = "column";
        container.style.gap = "0.5rem";
        document.body.appendChild(container);
      }

      const toast = document.createElement("div");
      const tone = (type || "info").toLowerCase();
      const colors = {
        success: { bg: "#ecfdf5", fg: "#166534", bd: "#86efac" },
        warning: { bg: "#fffbeb", fg: "#92400e", bd: "#fde68a" },
        error: { bg: "#fef2f2", fg: "#991b1b", bd: "#fecaca" },
        info: { bg: "#eff6ff", fg: "#1e3a8a", bd: "#bfdbfe" }
      };
      const c = colors[tone] || colors.info;
      toast.style.background = c.bg;
      toast.style.color = c.fg;
      toast.style.border = "1px solid " + c.bd;
      toast.style.borderRadius = "8px";
      toast.style.padding = "0.55rem 0.75rem";
      toast.style.fontSize = "0.8125rem";
      toast.style.boxShadow = "0 6px 18px rgba(0,0,0,0.12)";
      toast.style.maxWidth = "320px";
      toast.style.wordBreak = "break-word";
      toast.textContent = String(message || "");
      container.appendChild(toast);

      setTimeout(function() {
        toast.style.opacity = "0";
        toast.style.transition = "opacity .25s ease";
        setTimeout(function() {
          if (toast.parentNode) toast.parentNode.removeChild(toast);
        }, 260);
      }, 2800);
    } catch (e) {
      console.log(type ? "[" + type + "]" : "[info]", message);
    }
  };
}

const PAGE_ICONS = {
  "access-control.html": "ti ti-shield-lock"
};

async function ensureCurrentUserContext() {
  try {
    if (typeof currentRole === "string" && currentRole.trim()) {
      return {
        role: currentRole.trim().toUpperCase(),
        username: (typeof currentUser === "string" ? currentUser : "")
      };
    }
    if (typeof window.currentRole === "string" && window.currentRole.trim()) {
      return {
        role: window.currentRole.trim().toUpperCase(),
        username: window.currentUser || ""
      };
    }
    const r = await fetch("/api/auth/status");
    const d = await r.json();
    if (d && d.authenticated) {
      window.currentUser = d.displayName || d.username || "";
      window.currentRole = (d.role || "user").toUpperCase();
      return { role: window.currentRole, username: window.currentUser };
    }
  } catch (e) {
    console.error("Failed to load auth status for access-control", e);
  }
  return { role: "", username: "" };
}

function pageLabel(page) {
  if (!page) return "";
  const raw = String(page).replace(/^\/pages\//, "").replace(/\.html$/i, "");
  const name = raw.split("/").pop() || raw;
  return name
    .replace(/[_-]/g, " ")
    .replace(/([a-z])([A-Z])/g, "$1 $2")
    .replace(/\b\w/g, (c) => c.toUpperCase());
}

function pageIcon(page) {
  if (PAGE_ICONS[page]) return PAGE_ICONS[page];
  const key = String(page || "");
  if (key.startsWith("security/")) return "ti ti-shield-check";
  if (key.startsWith("warehouse/")) return "ti ti-warehouse";
  if (key.startsWith("purchase/")) return "ti ti-shopping-cart";
  if (key.startsWith("production/")) return "ti ti-settings";
  if (key.startsWith("qualityAssurance/")) return "ti ti-clipboard-check";
  if (key.startsWith("sales/")) return "ti ti-shopping-bag";
  if (key.startsWith("dispatch/")) return "ti ti-truck";
  if (key.startsWith("masterData/")) return "ti ti-database";
  if (key.startsWith("Laboratory/")) return "ti ti-flask";
  if (key.startsWith("labConfig/")) return "ti ti-flask";
  if (key.startsWith("contacts/")) return "ti ti-address-book";
  return "ti ti-file";
}

function getAccessControlRoot() {
  return (
    document.getElementById("accessControlRoot") ||
    document.getElementById("page-content") ||
    document.body
  );
}

async function initializeAccessControl() {
  const ctx = await ensureCurrentUserContext();
  if (ctx.role !== "ADMIN") {
    const root = getAccessControlRoot();
    root.innerHTML = '<div style="display:flex;flex-direction:column;align-items:center;padding:4rem;text-align:center;"><i class="ti ti-shield-lock" style="font-size:3rem;color:#ef4444;margin-bottom:1rem;"></i><h2 style="color:#1e293b;">Admin Only</h2><p style="color:#64748b;">This page is restricted to administrators.</p></div>';
    return;
  }
  await loadAllAcData();
  renderAll();
}
window["initializeAccess-control"] = initializeAccessControl;

async function loadAllAcData() {
  await loadAccessMatrix();
  await loadRolesInfo();
  await loadUsers();
}

async function fetchJsonNoCache(path) {
  const r = await fetch(apiNoCacheUrl(path), { cache: "no-store" });
  if (!r.ok) throw new Error(path + " returned " + r.status);
  return r.json();
}

function mergeRoleListFromRolesInfo() {
  const mergedRoles = new Set(acRoles);
  for (const role of Object.keys(acRolesInfo || {})) mergedRoles.add(role);
  acRoles = Array.from(mergedRoles).sort();
}

async function syncUsersSection(attempts = 8, waitMs = 350) {
  for (let i = 0; i < attempts; i++) {
    try {
      const users = await fetchJsonNoCache("/api/access/users");
      const rolesInfoResp = await fetchJsonNoCache("/api/access/roles");

      acUsers = Array.isArray(users) ? users : [];
      acRolesInfo = rolesInfoResp.roles || {};
      acPages = rolesInfoResp.allPages || acPages;
      mergeRoleListFromRolesInfo();

      renderUsersTable();
      renderRoleMembers();
      updateStats();
      return true;
    } catch (e) {
      if (i === attempts - 1) {
        console.error("syncUsersSection failed", e);
        return false;
      }
      await delay(waitMs);
    }
  }
  return false;
}

async function syncUsersSectionAfterCreate(createdUsername, attempts = 10, waitMs = 500) {
  const target = String(createdUsername || "").toLowerCase();
  for (let i = 0; i < attempts; i++) {
    try {
      const users = await fetchJsonNoCache("/api/access/users");
      const hasCreated = Array.isArray(users) && users.some((u) => String(u.username || "").toLowerCase() === target);
      if (!hasCreated) {
        await delay(waitMs);
        continue;
      }

      const rolesInfoResp = await fetchJsonNoCache("/api/access/roles");
      acUsers = users;
      acRolesInfo = rolesInfoResp.roles || {};
      acPages = rolesInfoResp.allPages || acPages;
      mergeRoleListFromRolesInfo();

      renderUsersTable();
      renderRoleMembers();
      updateStats();
      return true;
    } catch (e) {
      if (i === attempts - 1) {
        console.error("syncUsersSectionAfterCreate failed", e);
        return false;
      }
      await delay(waitMs);
    }
  }
  return false;
}

async function syncRolesSection(attempts = 8, waitMs = 350) {
  for (let i = 0; i < attempts; i++) {
    try {
      const accessResp = await fetchJsonNoCache("/api/access");
      const rolesInfoResp = await fetchJsonNoCache("/api/access/roles");

      acPages = accessResp.pages || rolesInfoResp.allPages || acPages;
      acMatrix = accessResp.matrix || {};
      acRolesInfo = rolesInfoResp.roles || {};
      mergeRoleListFromRolesInfo();

      renderRoleMembers();
      renderMatrix();
      updateStats();
      return true;
    } catch (e) {
      if (i === attempts - 1) {
        console.error("syncRolesSection failed", e);
        return false;
      }
      await delay(waitMs);
    }
  }
  return false;
}

function apiNoCacheUrl(path) {
  const sep = path.includes("?") ? "&" : "?";
  return path + sep + "_ts=" + Date.now() + "_r=" + Math.floor(Math.random() * 1000000);
}

async function loadAccessMatrix() {
  try {
    const r = await fetch(apiNoCacheUrl("/api/access"), { cache: "no-store" });
    const d = await r.json();
    acPages = d.pages || acPages;
    acMatrix = d.matrix || {};
  } catch(e) { console.error(e); }
}
async function loadRolesInfo() {
  try {
    const r = await fetch(apiNoCacheUrl("/api/access/roles"), { cache: "no-store" });
    const d = await r.json();
    acRolesInfo = d.roles || {};
    acPages = d.allPages || acPages;

    const mergedRoles = new Set(acRoles);
    for (const role of Object.keys(acRolesInfo)) mergedRoles.add(role);
    acRoles = Array.from(mergedRoles).sort();
  } catch(e) { console.error(e); }
}
async function loadUsers() {
  try { const r = await fetch(apiNoCacheUrl("/api/access/users"), { cache: "no-store" }); if (r.ok) acUsers = await r.json(); } catch(e) { console.error(e); }
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function refreshAllWithRetry(validateFn, attempts = 10, waitMs = 500) {
  for (let i = 0; i < attempts; i++) {
    await loadAllAcData();
    if (!validateFn || validateFn()) return true;
    await delay(waitMs);
  }
  return false;
}

function renderAll() { renderUsersTable(); renderRoleMembers(); renderMatrix(); updateStats(); }

/* ═══════════════════════════════════════
   USER TABLE
   ═══════════════════════════════════════ */

function renderUsersTable() {
  const tbody = document.getElementById("acUsersTable");
  if (!tbody) return;
  if (acUsers.length === 0) { tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;color:#94a3b8;padding:1.5rem;">No users</td></tr>'; return; }

  tbody.innerHTML = acUsers.map(u => {
    const cls = (u.role||'').toLowerCase();
    const enabled = u.enabled !== false;
    return `<tr style="${enabled ? '' : 'opacity:0.55;'}">
      <td style="padding:.6rem 1rem;"><strong style="font-size:.85rem;">${esc(u.username)}</strong></td>
      <td style="font-size:.85rem;">${esc(u.displayName)}</td>
      <td><span class="role-badge ${cls}"><i class="${u.role==='ADMIN'?'ti ti-crown':'ti ti-user'}"></i> ${esc(u.role)}</span></td>
      <td><span class="ac-status ${enabled?'active':'disabled'}">${enabled?'Active':'Disabled'}</span></td>
      <td>
        <div style="display:flex;gap:4px;">
          <button class="ac-btn" onclick="openEditUserModal(${u.id})" title="Edit"><i class="ti ti-edit"></i></button>
          <button class="ac-btn" onclick="openResetPwModal(${u.id},'${esc(u.username)}')" title="Reset Password"><i class="ti ti-key"></i></button>
          <button class="ac-btn" onclick="toggleUserEnabled(${u.id})" title="${enabled?'Disable':'Enable'}"><i class="${enabled?'ti ti-ban':'ti ti-check'}"></i></button>
          <button class="ac-btn danger" onclick="deleteUser(${u.id},'${esc(u.username)}')" title="Delete"><i class="ti ti-trash"></i></button>
        </div>
      </td>
    </tr>`;
  }).join("");
}

/* ═══════════════════════════════════════
   CREATE USER
   ═══════════════════════════════════════ */

function openCreateUserModal() {
  document.getElementById("newUserUsername").value = "";
  document.getElementById("newUserDisplayName").value = "";
  document.getElementById("newUserPassword").value = "";
  hideError("createUserError");
  populateRoleSelect("newUserRole");
  openModal("acCreateUserModal");
}

async function createUser() {
  const username = document.getElementById("newUserUsername").value.trim();
  const displayName = document.getElementById("newUserDisplayName").value.trim();
  const password = document.getElementById("newUserPassword").value;
  const role = document.getElementById("newUserRole").value;
  if (!username||!displayName||!password||!role) { showError("createUserError","All fields are required"); return; }
  if (username.length < 3) { showError("createUserError","Username must be at least 3 characters"); return; }
  if (password.length < 6) { showError("createUserError","Password must be at least 6 characters"); return; }

  try {
    const res = await fetch("/api/access/users", { method:"POST", headers:{"Content-Type":"application/json"},
      body:JSON.stringify({username,displayName,password,role}) });
    const data = await res.json();
    if (res.ok && data.success) {
      const tempId = Number(data.id || data.userId || Date.now());
      const exists = acUsers.some((u) => String(u.username || "").toLowerCase() === username.toLowerCase());
      if (!exists) {
        acUsers.unshift({
          id: tempId,
          username,
          displayName,
          role,
          enabled: true
        });
      }
      if (!acRolesInfo[role]) acRolesInfo[role] = [];
      if (!acRolesInfo[role].some((u) => String(u.username || "").toLowerCase() === username.toLowerCase())) {
        acRolesInfo[role].push({ username, displayName });
      }
      closeModal("acCreateUserModal");
      showNotification("User '" + username + "' created", "success");
      renderUsersTable();
      renderRoleMembers();
      updateStats();

      // Force one immediate API refresh after create, then poll if backend is still lagging.
      await loadUsers();
      const hasCreatedAfterImmediateRefresh = acUsers.some((u) => String(u.username || "").toLowerCase() === username.toLowerCase());
      if (hasCreatedAfterImmediateRefresh) {
        await loadRolesInfo();
        renderUsersTable();
        renderRoleMembers();
        updateStats();
      } else {
        const synced = await syncUsersSectionAfterCreate(username);
        if (!synced) showNotification("User created. Server list is still updating; keeping local view for now.", "warning");
      }
    } else { showError("createUserError", data.error || "Failed to create user"); }
  } catch(e) { showError("createUserError", e.message); }
}

/* ═══════════════════════════════════════
   EDIT USER
   ═══════════════════════════════════════ */

function openEditUserModal(userId) {
  const user = acUsers.find(u => u.id === userId);
  if (!user) return;
  document.getElementById("editUserId").value = userId;
  document.getElementById("editUserUsername").value = user.username;
  document.getElementById("editUserDisplayName").value = user.displayName;
  hideError("editUserError");
  populateRoleSelect("editUserRole", user.role);
  openModal("acEditUserModal");
}

async function saveEditUser() {
  const id = document.getElementById("editUserId").value;
  const displayName = document.getElementById("editUserDisplayName").value.trim();
  const role = document.getElementById("editUserRole").value;
  if (!displayName) { showError("editUserError","Display name is required"); return; }

  try {
    const res = await fetch("/api/access/users/" + id, { method:"PUT", headers:{"Content-Type":"application/json"},
      body:JSON.stringify({displayName,role}) });
    const data = await res.json();
    if (res.ok && data.success) {
      const user = acUsers.find((u) => String(u.id) === String(id));
      if (user) {
        user.displayName = displayName;
        user.role = role;

        for (const roleKey of Object.keys(acRolesInfo || {})) {
          acRolesInfo[roleKey] = (acRolesInfo[roleKey] || []).filter(
            (member) => String(member.username || "").toLowerCase() !== String(user.username || "").toLowerCase()
          );
        }
        if (!acRolesInfo[role]) acRolesInfo[role] = [];
        acRolesInfo[role].push({ username: user.username, displayName });
      }

      closeModal("acEditUserModal");
      showNotification("User updated", "success");
      renderUsersTable();
      renderRoleMembers();
      updateStats();

      const synced = await syncUsersSection();
      if (!synced) showNotification("User updated. Latest list is still syncing from server.", "warning");
    } else { showError("editUserError", data.error || "Failed to update"); }
  } catch(e) { showError("editUserError", e.message); }
}

/* ═══════════════════════════════════════
   RESET PASSWORD
   ═══════════════════════════════════════ */

function openResetPwModal(userId, username) {
  document.getElementById("resetPwUserId").value = userId;
  document.getElementById("resetPwUsername").textContent = username;
  document.getElementById("resetPwNewPassword").value = "";
  hideError("resetPwError");
  openModal("acResetPwModal");
}

async function resetPassword() {
  const id = document.getElementById("resetPwUserId").value;
  const password = document.getElementById("resetPwNewPassword").value;
  if (!password || password.length < 6) { showError("resetPwError","Password must be at least 6 characters"); return; }

  try {
    const res = await fetch("/api/access/users/" + id + "/reset-password", { method:"PUT", headers:{"Content-Type":"application/json"},
      body:JSON.stringify({password}) });
    const data = await res.json();
    if (res.ok && data.success) {
      closeModal("acResetPwModal");
      showNotification("Password reset successfully", "success");
    } else { showError("resetPwError", data.error || "Failed"); }
  } catch(e) { showError("resetPwError", e.message); }
}

/* ═══════════════════════════════════════
   TOGGLE / DELETE USER
   ═══════════════════════════════════════ */

async function toggleUserEnabled(userId) {
  try {
    const res = await fetch("/api/access/users/" + userId + "/toggle", { method:"PUT" });
    const data = await res.json();
    if (res.ok) { showNotification(data.username + " is now " + (data.enabled?"active":"disabled"), data.enabled?"success":"warning"); await loadUsers(); renderUsersTable(); updateStats(); }
    else { showNotification(data.error || "Failed", "error"); }
  } catch(e) { showNotification(e.message, "error"); }
}

async function deleteUser(userId, username) {
  if (!confirm("Permanently delete user '" + username + "'?\n\nThis cannot be undone.")) return;
  try {
    const res = await fetch("/api/access/users/" + userId, { method:"DELETE" });
    const data = await res.json();
    if (res.ok && data.success) {
      acUsers = acUsers.filter((u) => String(u.id) !== String(userId));
      for (const role of Object.keys(acRolesInfo || {})) {
        acRolesInfo[role] = (acRolesInfo[role] || []).filter((u) => String(u.username || "").toLowerCase() !== String(username || "").toLowerCase());
      }
      renderUsersTable();
      renderRoleMembers();
      updateStats();
      showNotification("User '" + username + "' deleted", "success");
      const synced = await syncUsersSection();
      if (!synced) showNotification("User deleted, but latest list could not be fetched yet", "warning");
    }
    else { showNotification(data.error || "Failed", "error"); }
  } catch(e) { showNotification(e.message, "error"); }
}

/* ═══════════════════════════════════════
   CREATE / DELETE ROLE
   ═══════════════════════════════════════ */

function openCreateRoleModal() {
  document.getElementById("newRoleName").value = "";
  hideError("createRoleError");
  const sel = document.getElementById("newRoleCopyFrom");
  sel.innerHTML = '<option value="">— Start with no access —</option>';
  acRoles.forEach(r => { sel.innerHTML += '<option value="' + r + '">' + r + '</option>'; });
  openModal("acCreateRoleModal");
}

async function createRole() {
  const roleName = document.getElementById("newRoleName").value.toUpperCase().replace(/[^A-Z0-9_]/g,'').trim();
  const copyFrom = document.getElementById("newRoleCopyFrom").value;
  if (!roleName || roleName.length < 2) { showError("createRoleError","Role name must be at least 2 characters (letters/digits/underscores)"); return; }

  try {
    const res = await fetch("/api/access/roles", { method:"POST", headers:{"Content-Type":"application/json"},
      body:JSON.stringify({roleName, copyFrom}) });
    const data = await res.json();
    if (res.ok && data.success) {
      if (!acRoles.includes(roleName)) acRoles.push(roleName);
      acRoles.sort();
      if (!acRolesInfo[roleName]) acRolesInfo[roleName] = [];
      if (!acMatrix[roleName]) acMatrix[roleName] = {};
      closeModal("acCreateRoleModal");
      renderRoleMembers();
      renderMatrix();
      updateStats();
      showNotification("Role '" + roleName + "' created", "success");
      const synced = await syncRolesSection();
      if (!synced) showNotification("Role created, but latest role data could not be fetched yet", "warning");
    } else { showError("createRoleError", data.error || "Failed"); }
  } catch(e) { showError("createRoleError", e.message); }
}

async function deleteRole(roleName) {
  if (!confirm("Delete role '" + roleName + "'?\n\nAll page-access entries for this role will be removed.\nUsers must be reassigned first.")) return;
  try {
    const res = await fetch("/api/access/roles/" + encodeURIComponent(roleName), { method:"DELETE" });
    const data = await res.json();
    if (res.ok && data.success) {
      acRoles = acRoles.filter((r) => r !== roleName);
      delete acRolesInfo[roleName];
      delete acMatrix[roleName];
      renderRoleMembers();
      renderMatrix();
      updateStats();
      showNotification("Role '" + roleName + "' deleted", "success");
      const synced = await syncRolesSection();
      if (!synced) showNotification("Role deleted, but latest role data could not be fetched yet", "warning");
    }
    else { showNotification(data.error || "Failed to delete role", "error"); }
  } catch(e) { showNotification(e.message, "error"); }
}

/* ═══════════════════════════════════════
   ROLES & MEMBERS
   ═══════════════════════════════════════ */

function renderRoleMembers() {
  const container = document.getElementById("roleMembersContainer");
  if (!container) return;
  if (Object.keys(acRolesInfo).length === 0) { container.innerHTML = '<p style="color:#94a3b8;text-align:center;">No roles</p>'; return; }

  let html = '<div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:1rem;">';
  for (const [role, users] of Object.entries(acRolesInfo)) {
    const cls = role.toLowerCase();
    const allowedCount = acMatrix[role] ? Object.values(acMatrix[role]).filter(v=>v===true).length : (role==="ADMIN"?acPages.length:0);
    const canDelete = role !== "ADMIN" && users.length === 0;

    html += `<div style="background:#fafbfc;border:1px solid #e2e8f0;border-radius:10px;padding:1rem;">
      <h4 style="margin:0 0 .5rem;font-size:.875rem;display:flex;align-items:center;gap:.5rem;">
        <span class="role-badge ${cls}"><i class="${role==='ADMIN'?'ti ti-crown':'ti ti-user'}"></i> ${role}</span>
        <span style="font-weight:400;color:#94a3b8;font-size:.75rem;">${users.length} user${users.length!==1?'s':''} · ${allowedCount}/${acPages.length} pages</span>
        ${canDelete ? `<button onclick="deleteRole('${role}')" style="margin-left:auto;background:none;border:none;color:#ef4444;cursor:pointer;font-size:1rem;" title="Delete role"><i class="ti ti-trash"></i></button>` : ''}
      </h4>
      <div style="display:flex;flex-wrap:wrap;">
        ${users.map(u => `<span class="role-member" title="${u.username}"><i class="ti ti-user"></i>${esc(u.displayName)}<span style="color:#94a3b8;font-size:.7rem;">(${u.username})</span></span>`).join("")}
        ${users.length===0 ? '<span style="color:#cbd5e1;font-size:.8rem;">No users assigned</span>' : ''}
      </div>
    </div>`;
  }
  html += '</div>';
  container.innerHTML = html;
}

/* ═══════════════════════════════════════
   PERMISSION MATRIX
   ═══════════════════════════════════════ */

function renderMatrix() {
  const table = document.getElementById("accessMatrixTable");
  if (!table) return;
  const editableRoles = acRoles.filter(r => r !== "ADMIN");

  let hdr = '<tr><th style="text-align:left;min-width:180px;padding:.6rem 1rem;">Page</th>';
  hdr += '<th><span class="role-badge admin"><i class="ti ti-crown"></i> ADMIN</span></th>';
  for (const role of editableRoles) { hdr += `<th><span class="role-badge ${role.toLowerCase()}"><i class="ti ti-user"></i> ${role}</span></th>`; }
  hdr += '</tr>';

  let body = '';
  for (const page of acPages) {
    const label = pageLabel(page);
    body += `<tr><td style="padding:.5rem 1rem;"><div class="page-name">${label}</div></td>`;
    body += '<td><input type="checkbox" checked disabled title="Admin always has full access"></td>';
    for (const role of editableRoles) {
      const allowed = acMatrix[role]&&acMatrix[role][page]===true;
      body += `<td><input type="checkbox" ${allowed?"checked":""} onchange="toggleAccess('${role}','${page}',this.checked)" title="${role} → ${label}"></td>`;
    }
    body += '</tr>';
  }

  table.className = "table table-hover align-middle mb-0 ac-matrix";
  table.innerHTML = `<thead class="ac-thead">${hdr}</thead><tbody>${body}</tbody>`;
}

async function toggleAccess(role, page, allowed) {
  showSaveIndicator(true);
  try {
    const res = await fetch("/api/access", { method:"PUT", headers:{"Content-Type":"application/json"}, body:JSON.stringify({role,page,allowed}) });
    if (res.ok) { if(!acMatrix[role])acMatrix[role]={}; acMatrix[role][page]=allowed; showNotification(`${role} ${allowed?"granted":"revoked"} → ${pageLabel(page)}`,allowed?"success":"warning"); }
    else { const e=await res.json(); showNotification(e.error||"Failed","error"); renderMatrix(); }
  } catch(e) { showNotification(e.message,"error"); renderMatrix(); }
  finally { showSaveIndicator(false); }
}

/* ═══════════════════════════════════════
   STATS
   ═══════════════════════════════════════ */

function updateStats() {
  const el = (id,v) => { const e=document.getElementById(id); if(e) e.textContent=v; };
  el("acTotalRoles", acRoles.length);
  el("acTotalPages", acPages.length);
  el("acTotalUsers", acUsers.length);
  el("acActiveUsers", acUsers.filter(u=>u.enabled!==false).length);
}

/* ═══════════════════════════════════════
   HELPERS
   ═══════════════════════════════════════ */

function populateRoleSelect(selectId, selected) {
  const sel = document.getElementById(selectId);
  if (!sel) return;
  sel.innerHTML = "";
  acRoles.forEach(r => { sel.innerHTML += `<option value="${r}" ${r===selected?'selected':''}>${r}</option>`; });
}

function openModal(id) {
  const m = document.getElementById(id);
  if (!m) return;
  if (window.jQuery && typeof window.jQuery.fn.modal === "function") {
    window.jQuery(m).modal("show");
    return;
  }
  if (window.bootstrap && window.bootstrap.Modal) {
    window.bootstrap.Modal.getOrCreateInstance(m).show();
    return;
  }
  m.style.display = "block";
  m.classList.add("show");
}
function closeModal(id) {
  const m = document.getElementById(id);
  if (!m) return;
  if (window.jQuery && typeof window.jQuery.fn.modal === "function") {
    window.jQuery(m).modal("hide");
    return;
  }
  if (window.bootstrap && window.bootstrap.Modal) {
    const instance = window.bootstrap.Modal.getInstance(m) || window.bootstrap.Modal.getOrCreateInstance(m);
    instance.hide();
    return;
  }
  m.classList.remove("show");
  m.style.display = "none";
}
function showError(id,msg) { const e=document.getElementById(id); if(e){e.textContent=msg;e.style.display="block";} }
function hideError(id) { const e=document.getElementById(id); if(e)e.style.display="none"; }
function showSaveIndicator(show) { const e=document.getElementById("acSaveIndicator"); if(e)e.style.display=show?"block":"none"; }
function esc(s) { if(!s)return""; const d=document.createElement("div"); d.textContent=String(s); return d.innerHTML; }