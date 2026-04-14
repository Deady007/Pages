let productRequestList = [];
let currentRequestId = '';
let currentRequestItems = [];
let currentRequestRecord = null;

$(document).ready(function () {
    fetchProductRequests();
});

function fetchProductRequests() {
    const payload = {
        fields: 'Id;Number;ProcessStatus;Status;Type;Remarks;CreatedDate;RequestedBy.Id;RequestedByName;IssuedBy.Id;IssuedByName',
        listOrderBy: 'CreatedDate DESC'
    };
    JSUTIL.callAJAXPost('/data/ProductRequest/list', JSON.stringify(payload), function (res) {
        const data = Array.isArray(res?.data) ? res.data : (res || []);
        productRequestList = data;
        renderRequestTable(productRequestList);
        if (currentRequestId) {
            const exists = productRequestList.find(r => String(r.id) === String(currentRequestId));
            if (!exists) {
                clearDetail();
            }
        }
    });
}

function renderRequestTable(list) {
    const body = $('#productIssueListTable');
    body.empty();
    if (!list.length) {
        body.html('<tr><td colspan="6" class="text-center text-muted">No product requests found.</td></tr>');
        return;
    }
    list.forEach(function (item) {
        const status = getStatusValue(item);
        const statusLabel = buildStatusLabel(status);
        const created = item.createdDate || item.CreatedDate || '';
        const canAct = canIssueStatus(status);
        const requestedByName = item.requestedByName || item.RequestedByName || item.requestedBy?.name || item.RequestedBy?.Name || '';
        const issueButton = canAct
            ? `<button class="btn btn-success btn-xs issue-request" data-id="${item.id}">Issue</button>`
            : '';
        const rejectButton = canAct
            ? `<button class="btn btn-danger btn-xs reject-request" data-id="${item.id}">Reject</button>`
            : '';
        body.append(`
            <tr>
                <td><button class="btn btn-link view-request" data-id="${item.id}">${item.number || item.Id || ''}</button></td>
                <td>${item.remarks || ''}</td>
                <td>${requestedByName}</td>
                <td>${JSUTIL.getDate ? JSUTIL.getDate(created) : created}</td>
                <td>${statusLabel}</td>
                <td>
                    <button class="btn btn-success btn-xs view-request" data-id="${item.id}">View</button>
                    ${issueButton}
                    ${rejectButton}
                </td>
            </tr>
        `);
    });
}

$(document).on('click', '.view-request', function () {
    const id = $(this).data('id');
    selectRequest(id);
});

$(document).on('click', '.issue-request', function () {
    const id = $(this).data('id');
    selectRequest(id, function () {
        issueCurrentRequest();
    });
});

$(document).on('click', '.reject-request', function () {
    const id = $(this).data('id');
    selectRequest(id, function () {
        rejectCurrentRequest();
    });
});

$(document).on('click', '.issue-request-btn', function () {
    issueCurrentRequest();
});

$(document).on('click', '.reject-request-btn', function () {
    rejectCurrentRequest();
});

function selectRequest(id, afterItemsLoaded) {
    const request = productRequestList.find(r => String(r.id) === String(id));
    if (!request) {
        JSUTIL.buildErrorModal('Request not found.');
        return;
    }
    currentRequestId = id;
    currentRequestRecord = request;
    renderDetail(request);
    fetchRequestItems(id, afterItemsLoaded);
}

function renderDetail(req) {
    $('#productIssueNumber').text(req.number || req.id || '-');
    const status = getStatusValue(req) || '-';
    $('#productIssueStatus').html(buildStatusLabel(status));
    $('#productIssueWorkOrder').text(req.remarks || '');
    $('#productIssueRequestedBy').text(req.requestedByName || req.RequestedByName || req.requestedBy?.name || req.RequestedBy?.Name || '');
    $('#productIssueRemarks').text(req.remarks || '');
    toggleDetailActions(status);
}

function clearDetail() {
    currentRequestId = '';
    currentRequestRecord = null;
    currentRequestItems = [];
    $('#productIssueNumber').text('-');
    $('#productIssueStatus').text('-');
    $('#productIssueWorkOrder').text('-');
    $('#productIssueRequestedBy').text('-');
    $('#productIssueRemarks').text('-');
    $('#productIssueMaterialTable').html('<tr><td colspan="4" class="text-center text-muted">Select a request to view items.</td></tr>');
    toggleDetailActions(null);
}

function fetchRequestItems(requestId, callback) {
    const payload = {
        fields: 'Id;Product.Id;QuantityRequested;QuantityIssued;QuantityRemaining;QuantityRemainingUom;QuantityLoss;QuantityLossUom;QuantityIssuedUom;ProductRequest.Id',
        conditions: [`ProductRequest.Id='${requestId}'`],
        logic: '{0}'
    };
    JSUTIL.callAJAXPost('/data/ProductRequestItem/list', JSON.stringify(payload), function (res) {
        const data = Array.isArray(res?.data) ? res.data : (res || []);
        currentRequestItems = data;
        renderRequestItems(data);
        if (typeof callback === 'function') {
            callback();
        }
    });
}

function renderRequestItems(items) {
    const body = $('#productIssueMaterialTable');
    body.empty();
    if (!items.length) {
        body.html('<tr><td colspan="4" class="text-center text-muted">No items.</td></tr>');
        return;
    }
    const canAct = canIssueStatus(getStatusValue(currentRequestRecord));
    items.forEach(function (item) {
        const requestedQty = Number(item.quantityRequested || 0);
        const issuedQty = Number(item.quantityIssued || 0);
        const remainingQty = Number(item.quantityRemaining || Math.max(requestedQty - issuedQty, 0));
        body.append(`
            <tr data-id="${item.id}" data-requested="${requestedQty}" data-issued="${issuedQty}" data-product-id="${item.product?.id || item.product?.Id || ''}">
                <td>${item.product?.name || ''}</td>
                <td>${item.product?.uom || ''}</td>
                <td>${requestedQty}</td>
                <td>
                    <input type="number" class="form-control issue-qty-input" value="${remainingQty}" min="0" max="${requestedQty}" ${canAct ? '' : 'readonly'}>
                    <div class="text-muted small">Issued: ${issuedQty} | Remaining: ${remainingQty}</div>
                </td>
            </tr>
        `);
    });
    toggleDetailActions(getStatusValue(currentRequestRecord));
}

function issueCurrentRequest() {
    if (!currentRequestId) {
        JSUTIL.buildErrorModal('Select a request before issuing.');
        return;
    }
    if (!canIssueStatus(getStatusValue(currentRequestRecord))) {
        JSUTIL.buildErrorModal('This request is already processed.');
        return;
    }
    const itemsPayload = [];
    const productUsageMap = {};
    let invalid = false;
    let hasQty = false;
    $('#productIssueMaterialTable tr').each(function () {
        const $row = $(this);
        const itemId = $row.data('id');
        const requested = Number($row.data('requested')) || 0;
        const issuedSoFar = Number($row.data('issued')) || 0;
        const productId = Number($row.data('product-id')) || null;
        const toIssue = Number($row.find('.issue-qty-input').val() || 0);
        const remaining = Math.max(requested - issuedSoFar, 0);
        if (toIssue < 0 || toIssue > remaining) {
            invalid = true;
            return false;
        }
        if (toIssue > 0) {
            hasQty = true;
        }
        const newIssuedTotal = issuedSoFar + toIssue;
        itemsPayload.push({
            id: itemId,
            quantityIssued: newIssuedTotal,
            quantityRemaining: Math.max(requested - newIssuedTotal, 0),
            quantityRequested: requested
        });
        if (productId) {
            productUsageMap[productId] = (productUsageMap[productId] || 0) + toIssue;
        }
    });
    if (invalid) {
        JSUTIL.buildErrorModal('Issue quantity must be between 0 and remaining quantity.');
        return;
    }
    if (!itemsPayload.length) {
        JSUTIL.buildErrorModal('No items available to issue.');
        return;
    }
    if (!hasQty) {
        JSUTIL.buildErrorModal('Enter quantity to issue for at least one item.');
        return;
    }
    setActionLoading(true);
    JSUTIL.callAJAXPost('/data/ProductRequestItem/upsert_multiple',
        JSON.stringify(itemsPayload),
        function () {
            updateProductOnHand(productUsageMap, function () {
                updateRequestStatus('Issued', function () {
                    JSUTIL.buildErrorModal('Request issued.');
                    fetchProductRequests();
                    fetchRequestItems(currentRequestId);
                });
            });
        },
        function (err) {
            console.error('Failed to issue items', err);
            JSUTIL.buildErrorModal('Failed to issue request items.');
            setActionLoading(false);
        }
    );
}

function rejectCurrentRequest() {
    if (!currentRequestId) {
        JSUTIL.buildErrorModal('Select a request before rejecting.');
        return;
    }
    if (!canIssueStatus(getStatusValue(currentRequestRecord))) {
        JSUTIL.buildErrorModal('This request is already processed.');
        return;
    }
    if (!confirm('Reject this request?')) {
        return;
    }
    setActionLoading(true);
    updateRequestStatus('Rejected', function () {
        JSUTIL.buildErrorModal('Request rejected.');
        fetchProductRequests();
        fetchRequestItems(currentRequestId);
    });
}

function updateRequestStatus(statusValue, callback) {
    const payload = { processStatus: statusValue };
    if (statusValue === 'Issued' && typeof employeeId !== 'undefined') {
        payload.issuedBy = { id: Number(employeeId) };
        payload.issuedByName = typeof employeeName !== 'undefined' ? employeeName : '';
    }
    JSUTIL.callAJAXPost(`/data/ProductRequest/update/${currentRequestId}`,
        JSON.stringify(payload),
        function (res) {
            if (currentRequestRecord) {
                currentRequestRecord.processStatus = statusValue;
                if (payload.issuedBy) {
                    currentRequestRecord.issuedBy = payload.issuedBy;
                    currentRequestRecord.issuedByName = payload.issuedByName;
                }
            }
            toggleDetailActions(statusValue);
            if (typeof callback === 'function') {
                callback(res);
            }
            setActionLoading(false);
        },
        function (err) {
            console.error('Failed to update request status', err);
            JSUTIL.buildErrorModal('Could not update request status.');
            setActionLoading(false);
        }
    );
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

function toggleDetailActions(status) {
    const hasRequest = !!currentRequestId;
    const canAct = hasRequest && canIssueStatus(status);
    const $actionButtons = $('.issue-request-btn, .reject-request-btn');
    $actionButtons.toggle(canAct).prop('disabled', !canAct);
    $('#productIssueMaterialTable .issue-qty-input').prop('readonly', !canAct);
}

function setActionLoading(isLoading) {
    $('.issue-request-btn, .reject-request-btn').prop('disabled', isLoading);
    $('.issue-request').prop('disabled', isLoading);
    $('.reject-request').prop('disabled', isLoading);
}

function updateProductOnHand(usageMap, onSuccess) {
    const productIds = Object.keys(usageMap || {}).filter(function (k) {
        return usageMap[k] > 0;
    });
    if (!productIds.length) {
        if (typeof onSuccess === 'function') onSuccess();
        return;
    }
    const conditionString = productIds.map(id => `'${id}'`).join(',');
    const queryPayload = {
        fields: 'Id;QuantityOnHand;',
        conditions: [`Id IN (${conditionString})`],
        logic: '{0}'
    };
    JSUTIL.callAJAXPost('/data/Product/list',
        JSON.stringify(queryPayload),
        function (res) {
            const data = Array.isArray(res?.data) ? res.data : (res || []);
            const updates = data.map(function (prod) {
                const pid = prod.id || prod.Id;
                const current = Number(prod.quantityOnHand || prod.QuantityOnHand || 0);
                const reduceBy = Number(usageMap[pid] || 0);
                const newQty = Math.max(current - reduceBy, 0);
                return {
                    id: pid,
                    quantityOnHand: newQty
                };
            });
            JSUTIL.callAJAXPost('/data/Product/upsert_multiple',
                JSON.stringify(updates),
                function () {
                    if (typeof onSuccess === 'function') onSuccess();
                },
                function (err) {
                    console.error('Failed to update product on hand', err);
                    JSUTIL.buildErrorModal('Issued, but failed to update product stock.');
                    setActionLoading(false);
                }
            );
        },
        function (err) {
            console.error('Failed to fetch product stock', err);
            JSUTIL.buildErrorModal('Could not fetch product quantities.');
            setActionLoading(false);
        }
    );
}
