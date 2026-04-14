window.ProductIssueModule = window.ProductIssueModule || {};

(function (module) {
    const state = module.state = module.state || {};
    state.list = state.list || [];
    state.typeFilter = state.typeFilter || 'ALL';
    state.currentRequestId = state.currentRequestId || '';
    state.currentRecord = state.currentRecord || null;
    state.currentItems = state.currentItems || [];

    $(document).ready(function () {
        fetchProductRequests();
    });

    $(document).on('productIssue:refreshList', function () {
        fetchProductRequests();
    });

    $(document).on('click', '.product-issue-filter', function () {
        const filterValue = $(this).data('filter');
        $('.product-issue-filter').removeClass('btn-primary active').addClass('btn-default');
        $(this).addClass('btn-primary active').removeClass('btn-default');
        applyRequestTypeFilter(filterValue);
    });

    $(document).on('click', '#productIssueListTable .view-request', function () {
        const id = $(this).data('id');
        triggerSelection(id, null);
    });

    $(document).on('click', '#productIssueListTable .issue-request', function () {
        const id = $(this).data('id');
        triggerSelection(id, 'issue');
    });

    $(document).on('click', '#productIssueListTable .reject-request', function () {
        const id = $(this).data('id');
        triggerSelection(id, 'reject');
    });

    function fetchProductRequests() {
        const payload = {
            fields: 'Id;Number;ProcessStatus;Status;Type;Remarks;CreatedDate;BatchNumber;RequestedBy.Id;RequestedByName;IssuedBy.Id;IssuedByName',
            conditions : ["ProcessStatus != 'Issued'" , "ProcessStatus != 'Cancelled'"],
            logic : "{0} AND {1}",
            listOrderBy: 'Id DESC'
        };
        JSUTIL.callAJAXPost('/data/ProductRequest/list', JSON.stringify(payload), function (res) {
            const data = Array.isArray(res?.data) ? res.data : (res || []);
            recordList = data;
            // data.sort(function (a, b) {
            //     const aId = Number(a?.id || a?.Id || 0);
            //     const bId = Number(b?.id || b?.Id || 0);
            //     return bId - aId;
            // });
            requestList = data;
            state.list = data;
            applyRequestTypeFilter(state.typeFilter);
        });
    }

    function applyRequestTypeFilter(filterValue) {
        state.typeFilter = filterValue || 'ALL';
        const filtered = filterRequestListByType(state.list, state.typeFilter);
        renderRequestTable(filtered);
    }

    function filterRequestListByType(list, filterValue) {
        const sourceList = Array.isArray(list) ? list : (Array.isArray(recordList) ? recordList : []);
        if (!sourceList.length) {
            return [];
        }
        if (!filterValue || filterValue === 'ALL') {
            return sourceList;
        }
        const target = normalizeTypeValue(filterValue);
        return sourceList.filter(function (item) {
            const rawType = item.type || item.Type || '';
            if (!rawType) {
                return false;
            }
            if (Array.isArray(rawType)) {
                return rawType.some(function (t) {
                    return normalizeTypeValue(t) === target;
                });
            }
            const type = normalizeTypeValue(rawType);
            if (type === target) {
                return true;
            }
            const parts = type.split(/[,/|]|\s+and\s+|\s*&\s*/i).map(function (t) { return normalizeTypeValue(t); }).filter(Boolean);
            return parts.includes(target);
        });
    }

    function renderRequestTable(list) {
        const body = $('#productIssueListTable');
        const $table = $('#productIssueListTable').closest('table');
        const hasDataTable = $.fn && $.fn.DataTable && $.fn.DataTable.isDataTable($table);
        if (hasDataTable) {
            const dt = $table.DataTable();
            dt.clear();
            if (Array.isArray(list) && list.length) {
                const rows = list.map(function (item) {
                    return buildRequestRowCells(item);
                });
                dt.rows.add(rows);
            }
            dt.page.len(25).draw(false);
            return;
        }

        body.empty();
        if (!Array.isArray(list) || !list.length) {
            body.html('<tr><td colspan="7" class="text-center text-muted">No product requests found.</td></tr>');
            return;
        }
        list.forEach(function (item) {
            const cells = buildRequestRowCells(item);
            body.append(`
                <tr>
                    <td>${cells[0]}</td>
                    <td>${cells[1]}</td>
                    <td>${cells[2]}</td>
                    <td>${cells[3]}</td>
                    <td>${cells[4]}</td>
                    <td>${cells[5]}</td>
                </tr>
            `);
        });
        reinitializeRequestTable($table);
    }

    function triggerSelection(id, action) {
        if (!id) {
            return;
        }
        const record = state.list.find(function (r) {
            return String(r.id || r.Id) === String(id);
        });
        if (!record) {
            JSUTIL.buildErrorModal('Request not found.');
            return;
        }
        $(document).trigger('productIssue:selected', [record, action]);
    }


    function getStatusValue(rec) {
        const raw = (rec && (rec.processStatus ?? rec.status)) || '';
        if (typeof raw === 'boolean') {
            return raw ? 'Active' : 'Inactive';
        }
        return raw;
    }

    function buildStatusLabel(status) {
        if (!status) return '<span class="label label-default">-</span>';
        let cls = 'label';
        const s = status.toLowerCase();
        if (s === 'requested' || s === 'pending') cls += ' label-info';
        else if (s === 'issued') cls += ' label-success';
        else if (s === 'rejected') cls += ' label-danger';
        else cls += ' label-default';
        return `<span class="${cls}">${status}</span>`;
    }

    function canIssueStatus(status) {
        const s = (status || '').toLowerCase();
        return s === '' || s === 'requested' || s === 'pending';
    }

    function parseWorkOrderRemark(remarks) {
        if (!remarks) {
            return { displayText: '-', bpcrId: null };
        }
        const match = remarks.match(/\(ID\s*:\s*([^)]+)\)/i);
        const bpcrId = match ? match[1].trim() : null;
        const displayText = match ? remarks.replace(match[0], '').trim() : remarks;
        return {
            displayText: displayText || '-',
            bpcrId: bpcrId || null
        };
    }

    function reinitializeRequestTable($table) {
        if (!window.$ || !$.fn || !$.fn.DataTable) {
            return;
        }
        const $target = $table && $table.length ? $table : $('#productIssueListTable').closest('table');
        try {
            if (!$.fn.DataTable.isDataTable($target)) {
                $target.DataTable({
                    pageLength: 25
                });
            } else {
                $target.DataTable().page.len(25).draw(false);
            }
        } catch (err) {
            console.error('DataTable init failed', err);
        }
    }

    function buildRequestRowCells(item) {
        const status = getStatusValue(item);
        const statusLabel = buildStatusLabel(status);
        const created = item.createdDate || item.CreatedDate || '';
        const requestedByName = item.requestedByName || item.RequestedByName
            || item.requestedBy?.name || item.RequestedBy?.Name || '';
        const typeLabel = item.type || item.Type || '-';
        const link = `<a href="javascript:void(0);" class="view-request" data-id="${item.id || item.Id}">${item.number || item.Id || ''}</a>`;
        return [
            link,
            item.batchNumber || '',
            requestedByName,
            JSUTIL.getDate ? JSUTIL.getDate(created) : created,
            typeLabel,
            statusLabel
        ];
    }

    function normalizeTypeValue(value) {
        if (value === null || typeof value === 'undefined') {
            return '';
        }
        let text = value.toString().trim().toLowerCase();
        if (!text) {
            return '';
        }
        if (text === 'rm' || text === 'raw materials' || text === 'raw material') {
            return 'raw material';
        }
        if (text === 'pm' || text === 'packing material' || text === 'packaging materials' || text === 'packaging material') {
            return 'packaging material';
        }
        if (text === 'fg' || text === 'finished goods' || text === 'finished good') {
            return 'finished goods';
        }
        if (text === 'ig' || text === 'intermediate goods' || text === 'intermediate good') {
            return 'intermediate goods';
        }
        if (text === 'gm' || text === 'general materials' || text === 'general material') {
            return 'general material';
        }
        if (text === 'mix' || text === 'mixed') {
            return 'mixed';
        }
        return text;
    }

    module.getStatusValue = getStatusValue;
    module.buildStatusLabel = buildStatusLabel;
    module.canIssueStatus = canIssueStatus;
    module.parseWorkOrderRemark = parseWorkOrderRemark;
    module.getProductRequestList = function () {
        return state.list.slice();
    };
})(window.ProductIssueModule);
