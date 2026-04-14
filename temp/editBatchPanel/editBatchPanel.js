let currentBPCRId = '';
let currentBPCRRecord = null;
const BPCR_QA_STATUS = 'In Review';
const BPCR_APPROVED_STATUS = 'Approved';
const BPCR_REJECTED_STATUS = 'Rejected';
const BPCR_ASSIGNED_STATUS = 'Assigned';
const BPCR_START_PRODUCTION_STATUS = 'In-Production';
const BPCR_END_PRODUCTION_STATUS = 'Production Ended';
const BPCR_COMPLETED_STATUS = 'Under Test';
const BPCR_PRODUCTION_OUTPUT_STATUS = 'Production Output';
const BPCR_PACKING_LIST_STATUS = 'Packing List Creation';
const BPCR_RELEASE_STATUS = 'Release';
const TEST_REQUEST_SOURCE = 'BPCR';
const isBPCRReviewPage = typeof window !== 'undefined'
    && (window.location.pathname.indexOf('bprcReview') >= 0);
let materialEditEnabled = false;
let currentStatusValue = '';
let bpcrMaterialItems = [];
let bpcrIssuedQuantityMap = {};
let bpcrRequestedQuantityMap = {};
let issueRequestItems = [];
let hasExistingMaterialRequest = false;
let materialRequestCheckPending = false;
const requestedMaterialLots = new Set();
let finalProductItem = null;
let packingListEntries = [];
let packingListEditEnabled = false;
let packingListLoadedFromServer = false;
let packingListFetchInProgress = false;
let productTransferTransactions = [];
let intermediateTransferTransactions = [];
const intermediateAddedTxnIds = new Set();
let transferPlantOptions = [];
let headerEditEnabled = false;
let finalProductEditEnabled = false;
let currentBOMId = '';
let currentBOMMeta = null;
const FINAL_PRODUCT_TYPE = 'Final Product';
const FINAL_PRODUCT_IN_WAREHOUSE_TYPE = 'Under-Test';
const BSR_ORDER_MATERIAL_TYPE = 'In-Warehouse';
const BSR_PRODUCT_LOT = 'BSR';
let pendingTestRequestBasePayload = null;
let currentStageIsFinal = false;
const BPCR_LABEL_PRINTER_STORAGE_KEY = 'productIssueLabelPrinters';
const BPCR_LABEL_LAST_PRINTER_KEY = 'productIssueSelectedPrinter';
const MATERIAL_TYPE_FILTER_ALL = 'all';
const MATERIAL_TYPE_FILTER_OPTIONS = [
    // { value: 'inProcess', label: 'In-Process Material' }
];
const bpcrLabelState = {
    records: [],
    printerOptions: [],
    selectedPrinterId: '',
    issueDate: null,
    finalData: null
};

$(document).ready(function () {
    $(document).on('click', '.selectBPCRBtn', function () {
        const orderId = $(this).attr('id');
        $('.newBPCRButton').hide();
        if (!orderId) {
            console.warn('No BPCR id found on selected row');
            return;
        }
        currentBPCRId = orderId;
        materialRequestCheckPending = true;
        hasExistingMaterialRequest = false;
        packingListEntries = [];
        packingListEditEnabled = false;
        resetRequestedMaterialLots();
        const hasRecordList = (typeof recordList !== 'undefined') && Array.isArray(recordList);
        currentBPCRRecord = hasRecordList
            ? recordList.find(r => String(r.id) === String(orderId))
            : null;
        if (!currentBPCRRecord) {
            console.warn('BPCR record not found in recordList for id:', orderId);
        }

        populateBPCRHeader(currentBPCRRecord, orderId);
        $('#allBPCRPanel').hide();
        $('#editBPCRPanel').show();
        setMaterialEditMode(false);
        loadBPCRDetails(orderId);
        refreshCurrentBPCRRecord(orderId);
    });
    $('.edit-bpcr-material').on('click', function () {
        setMaterialEditMode(true);
    });
    $('.request-materials').on('click', function () {
        prepareMaterialRequest();
    });
    $('#bpcrMaterialTypeFilter').on('change', function () {
        applyMaterialTypeFilter($(this).val());
    });
    $('.header-edit-btn').on('click', function () {
        setHeaderEditMode(true);
    });
    $('.header-save-btn').on('click', function () {
        saveHeaderDetails();
    });
    $('.print-bpcr-report').on('click', function () {
        handlePrintReport();
    });
    $('.print-final-label').on('click', function () {
        handlePrintFinalLabel();
    });
    $('.send-to-qc').on('click', function () {
        sendToQC();
    });
    $(document).on('click', '.finish-production-btn' , function () {
        handleFinishProduction();
    });
    $(document).on('click', '.add-transfer-btn', function () {
        openAddTransferModal();
    });
    $(document).on('click', '.get-intermediate-btn', function () {
        openIntermediateModal();
    });
    $(document).on('click', '.recall-transfer-btn', function () {
        handleRecallTransfer($(this));
    });
    $(document).on('click', '.delete-transfer-btn', function () {
        handleDeleteTransfer($(this));
    });
    $(document).on('click', '.bpcr-intermediate-get', function () {
        handleIntermediateGet($(this).data('txnId'));
    });
    $('#bpcrTransferTypeSelect').on('change', function () {
        handleTransferTypeChange();
    });
    $('#bpcrTransferPlantSelect').on('change', function () {
        handleTransferPlantChange();
    });
    $('#bpcrTransferSaveBtn').on('click', function () {
        saveTransferTransaction();
    });
    $('#bpcrAddTransferModal').on('hidden.bs.modal', function () {
        resetTransferModal();
    });
    $('.confirm-qc-sample-quantity').on('click', function () {
        confirmSampleQuantityAndSend();
    });
    $('.save-used-qty').on('click', function () {
        saveUsedQuantities();
    });
    $(document).on('input', '.bpcr-used-qty-input', function () {
        const $row = $(this).closest('tr');
        updateReturnedQuantityForRow($row);
    });
    $('.save-bpcr-material, #bpcr-material .request-materials').on('click', function () {
        saveBpcrChanges();
    });
    $(document).on('click', '.final-product-edit-btn', function () {
        setFinalProductEditMode(true);
    });
    $(document).on('click', '.final-product-save-btn', function () {
        saveFinalProductDetails();
    });
    $(document).on('click', '.packing-list-edit-btn', function () {
        setPackingListEditMode(true);
    });
    $(document).on('click', '.packing-list-save-btn', function () {
        savePackingListDetails();
    });
    $(document).on('change', '#finalProductNumberOfContainers', function () {
        syncPackingListWithContainerCount();
    });
    $(document).on('change', '#finalProductExpiryType', function () {
        updateFinalProductDateLabel();
    });
    $(document).on('click', '.cancel-issue-request-btn', function () {
        handleIssueRequestCancel($(this));
    });
    $('.close-edit-bpcr').on('click', function () {
        $('#editBPCRPanel').hide();
        $('#allBPCRPanel').show();
        $('.newBPCRButton').show();
        setMaterialEditMode(false);
        resetBPCRHeader(true);
    });

    $('#bpcrPrintLabelConfirmBtn').on('click', function () {
        submitBpcrLabelPrint(false);
    });
    $('#bpcrDownloadZplBtn').on('click', function () {
        submitBpcrLabelPrint(true);
    });
    $('#bpcrAddPrinterBtn').on('click', function () {
        addBpcrManualPrinter();
    });
    $('#bpcrPrinterRefreshBtn').on('click', function () {
        loadBpcrLabelPrinterOptions(true);
    });
    $('#bpcrPrinterSelect').on('change', function () {
        handleBpcrPrinterSelection($(this).val());
    });
    $('#bpcrLabelSelectAll').on('change', function () {
        const checked = $(this).prop('checked');
        $('#bpcrLabelTableBody .bpcr-label-row-check').prop('checked', checked);
    });
    $('#bpcrLabelTableBody').on('input', '.bpcr-label-gross-input, .bpcr-label-tare-input', function () {
        const $row = $(this).closest('tr');
        const gross = Number($row.find('.bpcr-label-gross-input').val());
        const tare = Number($row.find('.bpcr-label-tare-input').val());
        if (!Number.isFinite(gross) || !Number.isFinite(tare)) {
            return;
        }
        const net = Math.max(gross - tare, 0);
        $row.find('.bpcr-label-net-input').val(net);
    });
    $('#bpcrLabelModal').on('hidden.bs.modal', function () {
        setBpcrLabelModalMessage('', false);
    });

    setMaterialEditMode(false);
    applyMaterialTypeFilter();
});
function populateBPCRHeader(record, fallbackId) {
    resetBPCRHeader(false);
    const displayNumber = record?.number || fallbackId || '';
    const displayName = record?.name || fallbackId || '';
    const productInfo = getBPCRProductInfo(record);
    const stageInfo = getBPCRStageInfo(record);
    $('#editBPCRNumber').text(displayNumber + '  >  ' + currentBPCRRecord.name);
    $('#bpcrNumberDisplay').text(displayName || '-');
    $('#bpcrBatchNameInput').val(displayName || '');
    $('#bpcrProductDisplay').text(productInfo.display || '-');
    $('#bpcrStageDisplay').text(stageInfo.display || '-');
    var recipe = record?.productRecipe || record?.ProductRecipe || {};
    var masterBatchRecord = recipe.masterBatchRecord
        || recipe.MasterBatchRecord
        || recipe.name
        || recipe.Name
        || '';
    setCurrentStageFinalFlag(record);
    var bomLabel = (recipe.name || '') + (masterBatchRecord ? ' > ' + masterBatchRecord : '');
    $('#bpcrBOMDisplay').text(bomLabel || '-');
    $('#docCodeDisplay').text(currentBPCRRecord.number|| '-')
    $('#docCodeDateDisplay').text(formatDateDDMMYYYY(currentBPCRRecord.createdDate) || '-')
    setBatchMetaFromOrder(record);
    loadBOMMetaForHeader(record);
    $('#bpcrPlantDisplay').text(getBPCRPlantName(record) || '-');
    $('#bpcrOwnerDisplay').text(record?.orderOwner?.name || 'Not Assigned');
    currentStatusValue = normalizeIssuedStatus(record?.orderStatus || '');
    const actualStartTimeValue = record?.actualStartTime || record?.ActualStartTime || '';
    const actualStopTimeValue = record?.actualEndTime || record?.ActualEndTime || record?.actualStopTime || record?.ActualStopTime || '';
    $('#bpcrActualStartInput').val(formatDateTimeInputValue(actualStartTimeValue));
    $('#bpcrActualEndInput').val(formatDateTimeInputValue(actualStopTimeValue));
    toggleQAButton(currentStatusValue);
    toggleReviewButtons(currentStatusValue);
    toggleOwnerVisibility(currentStatusValue);
    toggleEditAvailability(currentStatusValue);
    toggleAssignButton(currentStatusValue);
    toggleProductionButtons(currentStatusValue);
    toggleFinalProductTab(currentStatusValue);
    toggleQCButton(currentStatusValue);
    toggleActualTimeEditControls(currentStatusValue);
    toggleTabsVisibility(currentStatusValue);
    toggleProductTransferButtons(currentStatusValue);
    updateProgressSteps(currentStatusValue);
    finalProductEditEnabled = false;
    setHeaderEditMode(false);
}

function resetBPCRHeader(resetState) {
    if (resetState) {
        currentBPCRId = '';
        currentBPCRRecord = null;
        hasExistingMaterialRequest = false;
        materialRequestCheckPending = false;
        resetRequestedMaterialLots();
    }
    currentStatusValue = '';
    bpcrIssuedQuantityMap = {};
    bpcrRequestedQuantityMap = {};
    issueRequestItems = [];
    finalProductItem = null;
    finalProductEditEnabled = false;
    currentBOMId = '';
    currentBOMMeta = null;
    currentStageIsFinal = false;
    $('#editBPCRNumber').text('');
    $('#bpcrNumberDisplay').text('-');
    $('#bpcrBatchNameInput').val('');
    $('#bpcrProductDisplay').text('-');
    $('#bpcrStageDisplay').text('-');
    $('#bpcrBOMDisplay').text('-');
    $('#bpcrBatchSizeDisplay').text('-');
    $('#bpcrPlantDisplay').text('-');
    $('#bpcrOwnerDisplay').text('-');
    $('#bpcrActualStartInput').val('');
    $('#bpcrActualEndInput').val('');
    toggleQAButton(null);
    toggleReviewButtons(null);
    toggleOwnerVisibility(null);
    toggleEditAvailability(null);
    toggleAssignButton(null);
    toggleProductionButtons(null);
    toggleFinalProductTab(null);
    toggleQCButton(null);
    toggleActualTimeEditControls(null);
    toggleProductTransferButtons(null);
}

function normalizeLotName(value) {
    return normalizeStepName(value || '');
}

function resetRequestedMaterialLots() {
    requestedMaterialLots.clear();
}

function isCancelledRequestStatus(statusValue) {
    const normalized = (statusValue || '').toString().toLowerCase();
    return normalized === 'cancelled' || normalized === 'canceled';
}

function markRequestedLot(lotLabel) {
    const normalized = normalizeLotName(lotLabel);
    if (!normalized) {
        return;
    }
    requestedMaterialLots.add(normalized);
}

function isLotAlreadyRequested(lotLabel) {
    const normalized = normalizeLotName(lotLabel);
    if (!normalized) {
        return false;
    }
    return requestedMaterialLots.has(normalized);
}

function refreshRequestedLotsFromIssueRequests(list) {
    resetRequestedMaterialLots();
    (list || []).forEach(function (item) {
        if (isCancelledRequestStatus(item?.status)) {
            return;
        }
        if (item?.lotLabel) {
            markRequestedLot(item.lotLabel);
        }
    });
    toggleAssignButton(currentStatusValue);
}

function parseRequestedLotFromRemarks(remarks) {
    if (!remarks) {
        return '';
    }
    const match = remarks.match(/lot\s*[:\-]\s*([a-z0-9 ._#\/-]+)/i);
    if (match && match[1]) {
        return match[1].trim();
    }
    return '';
}

function inferLotFromProduct(productId) {
    if (!productId) {
        return '';
    }
    const lots = [];
    (bpcrMaterialItems || []).forEach(function (item) {
        const pid = item?.product?.id || item?.product?.Id || item?.productId || item?.ProductId;
        if (String(pid) !== String(productId)) {
            return;
        }
        const lotLabel = (item.productLot || item.ProductLot || '').toString().trim();
        if (lotLabel) {
            lots.push(lotLabel);
        }
    });
    const uniqueLots = Array.from(new Set(lots));
    return uniqueLots.length === 1 ? uniqueLots[0] : '';
}

function getSelectedLotInfo() {
    const $filter = $('#bpcrMaterialTypeFilter');
    if (!$filter.length) {
        return { label: '', normalized: '', isStepFilter: false };
    }
    const $selectedOption = $filter.find('option:selected');
    const value = ($selectedOption.val() || '').toString().trim();
    const filterType = ($selectedOption.data('filter-type') || '').toString();
    const label = ($selectedOption.text() || value).trim();
    if (filterType === 'step') {
        const normalized = normalizeLotName($selectedOption.data('step-name') || value);
        return { label: value || label, normalized: normalized, isStepFilter: true };
    }
    return { label: '', normalized: '', isStepFilter: false };
}

function resolveLotForRequest() {
    const selected = getSelectedLotInfo();
    if (selected.isStepFilter && selected.normalized) {
        return selected;
    }
    const lotLabels = getUniqueMaterialSteps(bpcrMaterialItems || []);
    if (lotLabels.length === 1) {
        const label = lotLabels[0];
        return { label: label, normalized: normalizeLotName(label), isStepFilter: true };
    }
    return {
        label: '',
        normalized: '',
        isStepFilter: false,
        requiresSelection: lotLabels.length > 1
    };
}

function toggleQAButton(status) {
    $('.submit-bpcr-qa, .approve-bpcr, .reject-bpcr').hide();
}

function toggleReviewButtons(status) {
    $('.approve-bpcr, .reject-bpcr').hide();
}

function submitBPCRToQA() {
    if (!currentBPCRId) {
        console.log('Select a BPCR before submitting to QA.');
        return;
    }
    const payload = { orderStatus: BPCR_QA_STATUS };
    JSUTIL.callAJAXPost(`/data/ProductionOrder/update/${currentBPCRId}`,
        JSON.stringify(payload),
        function (res) {
            console.log('Submitted to QA for approval.');
            if (currentBPCRRecord) {
                currentBPCRRecord.orderStatus = BPCR_QA_STATUS;
            }
            populateBPCRHeader(currentBPCRRecord, currentBPCRId);
        }
    );
}

function updateBPCRStatus(statusValue, extraPayload, options) {
    if (!currentBPCRId) {
        // JSUTIL.buildErrorModal'Select a BPCR before updating status.');
        return;
    }
    const payload = Object.assign({}, extraPayload || {});
    payload.orderStatus = statusValue;
    const shouldReload = !(options && options.reload === false);
    JSUTIL.callAJAXPost(`/data/ProductionOrder/update/${currentBPCRId}`,
        JSON.stringify(payload),
        function () {
            // JSUTIL.buildErrorModal`BPCR ${statusValue}.`);
            if (currentBPCRRecord) {
                currentBPCRRecord.orderStatus = statusValue;
                if (extraPayload && typeof extraPayload === 'object') {
                    Object.keys(extraPayload).forEach(function (key) {
                        currentBPCRRecord[key] = extraPayload[key];
                    });
                }
            }
            populateBPCRHeader(currentBPCRRecord, currentBPCRId);
            if (shouldReload) {
                location.reload();
            }
        }
    );
}

function buildFinalProductPayload() {
    const qty = Number($('#finalProductQty').val());
    const isFinalStage = isFinalStageForCurrentBatch();
    const numberOfContainers = isFinalStage ? Number($('#finalProductNumberOfContainers').val()) : null;
    const productId = $('#finalProductProductId').val();
    const existingId = $('#finalProductId').val();
    const manufacturingDate = $('#finalProductManufacturingDate').val();
    const expiryDate = $('#finalProductExpiryDate').val();
    const expiryType = normalizeExpiryType($('#finalProductExpiryType').val() || 'Expiry');
    const docNumber = ($('#finalProductDocNumber').val() || getBPCRDocNumber() || '').toString().trim();
    const netQtyRaw = $('#finalProductNetQtyPerContainer').val();
    const variationRaw = $('#finalProductVariationPerContainer').val();
    const gradeValue = ($('#finalProductGrade').val() || '').toString().trim();
    const netQtyPerContainer = netQtyRaw === '' ? null : Number(netQtyRaw);
    const variationPerContainer = variationRaw === '' ? null : Number(variationRaw);
    if (isProductionEndedStatus(currentStatusValue)) {
        return { valid: false, message: 'Production has ended. Final product details are locked.' };
    }
    if (!Number.isFinite(qty) || qty <= 0) {
        return { valid: false, message: 'Enter produced quantity for the final product.' };
    }
    if (isFinalStage && (!Number.isFinite(numberOfContainers) || numberOfContainers <= 0)) {
        return { valid: false, message: 'Enter the number of containers/bags for the final product.' };
    }
    if (!productId) {
        return { valid: false, message: 'Final product is missing. Please refresh and try again.' };
    }
    if (!manufacturingDate) {
        return { valid: false, message: 'Enter manufacturing date for the final product.' };
    }
    const mfgTime = new Date(manufacturingDate).getTime();
    const expiryTime = expiryDate ? new Date(expiryDate).getTime() : null;
    if (expiryDate && Number.isFinite(mfgTime) && Number.isFinite(expiryTime) && expiryTime < mfgTime) {
        return { valid: false, message: 'Expiry date cannot be before manufacturing date.' };
    }
    if (netQtyRaw !== '' && !Number.isFinite(netQtyPerContainer)) {
        return { valid: false, message: 'Enter a valid net quantity per container.' };
    }
    if (variationRaw !== '' && !Number.isFinite(variationPerContainer)) {
        return { valid: false, message: 'Enter a valid variation per container.' };
    }
    if (!docNumber) {
        return { valid: false, message: 'BPCR number is missing for doc number.' };
    }
    const payload = {
        type: FINAL_PRODUCT_TYPE,
        plannedQuantity: qty,
        materialStatus: 'Actual',
        remarks: currentBPCRId,
        product: { id: Number(productId) },
        manufacturingDate: manufacturingDate,
        expiryDate: expiryDate,
        expiryType: expiryType,
        docNumber: docNumber
    };
    if (isFinalStage && Number.isFinite(numberOfContainers) && numberOfContainers > 0) {
        payload.numberOfContainers = numberOfContainers;
    }
    payload.actualQuantity = netQtyRaw === '' ? null : netQtyPerContainer;
    payload.variation = variationRaw === '' ? null : variationPerContainer;
    payload.grade = gradeValue || null;
    if (existingId) {
        payload.id = existingId;
    }
    return { valid: true, payload: payload };
}

function normalizeExpiryType(value) {
    const normalized = (value || '').toString().trim().toLowerCase();
    if (normalized === 'retest') {
        return 'Retest';
    }
    return 'Expiry';
}

function buildWarehouseFinalProductPayload() {
    const baseItem = finalProductItem || {};
    const id = baseItem.id || $('#finalProductId').val();
    const productId = baseItem?.product?.id
        || baseItem?.product?.Id
        || $('#finalProductProductId').val()
        || '';
    if (!id || !productId) {
        return null;
    }
    const plannedQty = firstDefinedNumber(
        baseItem.plannedQuantity,
        $('#finalProductQty').val(),
        0
    );
    const containerCount = firstDefinedNumber(
        baseItem.numberOfContainers,
        $('#finalProductNumberOfContainers').val(),
        0
    );
    const manufacturingDate = baseItem.manufacturingDate
        || $('#finalProductManufacturingDate').val()
        || '';
    const expiryDate = baseItem.expiryDate
        || $('#finalProductExpiryDate').val()
        || '';
    const retestDate = baseItem.retestDate
        || baseItem.expiryDate
        || $('#finalProductExpiryDate').val()
        || '';
    const expiryType = normalizeExpiryType(
        baseItem.expiryType
        || $('#finalProductExpiryType').val()
        || 'Expiry'
    );
    const netQtyRaw = $('#finalProductNetQtyPerContainer').val();
    const variationRaw = $('#finalProductVariationPerContainer').val();
    const gradeValue = ($('#finalProductGrade').val() || baseItem.grade || '').toString().trim();
    const netQtyPerContainer = netQtyRaw === '' ? baseItem.actualQuantity : Number(netQtyRaw);
    const variationPerContainer = variationRaw === '' ? baseItem.variation : Number(variationRaw);
    // const retestDate = baseItem.retestDate
    //     || $('#finalProductExpiryDate').val()
    //     || '';
    const docNumber = (baseItem.docNumber
        || $('#finalProductDocNumber').val()
        || getBPCRDocNumber()
        || '').toString().trim();
    const productCodeValue = buildWarehouseProductCodeValue(baseItem);
    const payload = {
        id: id,
        type: FINAL_PRODUCT_IN_WAREHOUSE_TYPE,
        plannedQuantity: plannedQty,
        quantityToWarehouse: plannedQty,
        materialStatus: baseItem.materialStatus || 'Actual',
        remarks: baseItem.remarks || currentBPCRId,
        product: { id: Number(productId) }
    };
    if (manufacturingDate) {
        payload.manufacturingDate = manufacturingDate;
    }
    if (expiryDate) {
        payload.expiryDate = expiryDate;
    }
    payload.expiryType = expiryType;
    if (netQtyRaw !== '') {
        payload.actualQuantity = Number.isFinite(netQtyPerContainer) ? netQtyPerContainer : null;
    } else if (baseItem.actualQuantity !== undefined) {
        payload.actualQuantity = baseItem.actualQuantity;
    }
    if (variationRaw !== '') {
        payload.variation = Number.isFinite(variationPerContainer) ? variationPerContainer : null;
    } else if (baseItem.variation !== undefined) {
        payload.variation = baseItem.variation;
    }
    if (gradeValue) {
        payload.grade = gradeValue;
    }
    if (docNumber) {
        payload.docNumber = docNumber;
    }
    if (productCodeValue) {
        payload.productCode = productCodeValue;
    }
    if (containerCount > 0) {
        payload.numberOfContainers = containerCount;
    }
    return payload;
}

function saveFinalProduct(payload, onSuccess, onError) {
    if (!payload) {
        if (typeof onError === 'function') onError();
        return;
    }
    JSUTIL.callAJAXPost('/data/OrderMaterial/upsert_multiple',
        JSON.stringify([payload]),
        function (res) {
            const newId = Array.isArray(res) ? res[0]?.id : res?.id;
            finalProductItem = {
                ...(finalProductItem || {}),
                ...payload,
                id: newId || payload.id || ''
            };
            $('#finalProductId').val(finalProductItem.id || '');
            if (typeof onSuccess === 'function') {
                onSuccess(res);
            }
        },
        function (err) {
            console.error('Failed to save final product', err);
            if (typeof onError === 'function') {
                onError(err);
            }
        }
    );
}

function updateFinalProductTypeToWarehouse(onSuccess, onError) {
    const payload = buildWarehouseFinalProductPayload();
    if (!payload) {
        console.warn('Final product details missing. Cannot update type to Under Test.');
        if (typeof onError === 'function') {
            onError();
        }
        return;
    }
    saveFinalProduct(payload, onSuccess, onError);
}

function handleEndProduction() {
    if (!currentBPCRId) {
        return;
    }
    const result = buildFinalProductPayload();
    console.log("result : ", result)
    if (!result.valid) {
        JSUTIL.buildErrorModal(result.message || 'Enter final product details before ending production.');
        return;
    }
    const timestamp = new Date().toISOString();
    saveFinalProduct(result.payload, function () {
        updateBPCRStatus(BPCR_END_PRODUCTION_STATUS, { actualEndTime: timestamp });
    }, function () {
        console.log('Could not save final product details. Please try again.');
    });
}

function handleStartProduction() {
    if (!currentBPCRId) {
        return;
    }
    const timestamp = new Date().toISOString();
    updateBPCRStatus(BPCR_START_PRODUCTION_STATUS, { actualStartTime: timestamp });
}

function buildActualTimePayload() {
    const startRaw = $('#bpcrActualStartInput').val();
    const endRaw = $('#bpcrActualEndInput').val();
    const payload = {};
    const parsedStart = parseDateTimeLocalValue(startRaw);
    const parsedEnd = parseDateTimeLocalValue(endRaw);
    if (startRaw && !parsedStart) {
        return { valid: false, message: 'Enter a valid actual start time.' };
    }
    if (endRaw && !parsedEnd) {
        return { valid: false, message: 'Enter a valid actual stop time.' };
    }
    if (parsedStart) {
        payload.actualStartTime = parsedStart;
    }
    if (parsedEnd) {
        payload.actualEndTime = parsedEnd;
    }
    if (parsedStart && parsedEnd && new Date(parsedEnd).getTime() < new Date(parsedStart).getTime()) {
        return { valid: false, message: 'Actual stop time cannot be before actual start time.' };
    }
    if (!Object.keys(payload).length) {
        return { valid: false, message: 'Enter start or stop time to save.' };
    }
    return { valid: true, payload: payload };
}

function saveActualTimes(options) {
    const opts = Object.assign({
        skipIfEmpty: false,
        allowSkipOnLock: false
    }, options || {});
    if (!currentBPCRId) {
        return Promise.resolve(false);
    }
    const startRaw = $('#bpcrActualStartInput').val();
    const endRaw = $('#bpcrActualEndInput').val();
    const noTimeEntered = !startRaw && !endRaw;
    if (noTimeEntered && opts.skipIfEmpty) {
        return Promise.resolve(false);
    }
    if (!canEditActualTimes(currentStatusValue)) {
        if (opts.allowSkipOnLock) {
            return Promise.resolve(false);
        }
        showActualTimeError('Actual start and stop times are locked after sending to QC.');
        return Promise.reject(new Error('Actual time editing is locked.'));
    }
    const result = buildActualTimePayload();
    if (!result.valid) {
        showActualTimeError(result.message || 'Enter valid actual start/stop times.');
        return Promise.reject(new Error(result.message || 'Invalid actual times'));
    }
    return new Promise(function (resolve, reject) {
        JSUTIL.callAJAXPost(`/data/ProductionOrder/update/${currentBPCRId}`,
            JSON.stringify(result.payload),
            function () {
                if (currentBPCRRecord) {
                    if (result.payload.actualStartTime) {
                        currentBPCRRecord.actualStartTime = result.payload.actualStartTime;
                    }
                    if (result.payload.actualEndTime) {
                        currentBPCRRecord.actualEndTime = result.payload.actualEndTime;
                    }
                }
                populateBPCRHeader(currentBPCRRecord, currentBPCRId);
                resolve(true);
            },
            function (err) {
                console.error('Failed to update actual times', err);
                reject(err || new Error('Failed to update actual times'));
            }
        );
    });
}

function buildMaterialPlanPayload() {
    if (!materialEditEnabled) {
        return [];
    }
    const rows = $('#editBPCRPanel tbody tr');
    const payload = [];
    rows.each(function () {
        const id = $(this).data('id');
        const plannedQty = Number($(this).data('planned-qty'));
        if (!id || !Number.isFinite(plannedQty)) {
            return;
        }
        payload.push({
            id: id,
            plannedQuantity: plannedQty
        });
    });
    return payload;
}

function saveMaterialPlanUpdates(payload) {
    const materialPayload = Array.isArray(payload) ? payload : buildMaterialPlanPayload();
    if (!materialPayload.length) {
        return Promise.resolve(false);
    }
    return new Promise(function (resolve, reject) {
        JSUTIL.callAJAXPost('/data/OrderMaterial/upsert_multiple',
            JSON.stringify(materialPayload),
            function (res) {
                console.log("Materials updated:", res);
                resolve(true);
            },
            function (err) {
                console.error('Failed to update materials', err);
                reject(err || new Error('Failed to update materials'));
            }
        );
    });
}

function saveBpcrChanges() {
    const materialPayload = buildMaterialPlanPayload();
    saveActualTimes({ skipIfEmpty: true, allowSkipOnLock: true })
        .then(function () {
            return saveMaterialPlanUpdates(materialPayload);
        })
        .then(function () {
            setMaterialEditMode(false);
        })
        .catch(function (err) {
            console.error('Failed to save BPCR changes', err);
        });
}

function saveHeaderDetails() {
    if (!currentBPCRId) {
        return;
    }
    const payload = {};
    const nameValue = ($('#bpcrBatchNameInput').val() || '').toString().trim();
    const currentName = currentBPCRRecord?.name || currentBPCRRecord?.Name || '';
    if (nameValue && nameValue !== currentName) {
        payload.name = nameValue;
    }
    const startRaw = $('#bpcrActualStartInput').val();
    const endRaw = $('#bpcrActualEndInput').val();
    const parsedStart = startRaw ? parseDateTimeLocalValue(startRaw) : '';
    const parsedEnd = endRaw ? parseDateTimeLocalValue(endRaw) : '';
    if (startRaw && !parsedStart) {
        showActualTimeError('Enter a valid actual start time.');
        return;
    }
    if (endRaw && !parsedEnd) {
        showActualTimeError('Enter a valid actual end time.');
        return;
    }
    if (parsedStart) {
        payload.actualStartTime = parsedStart;
    }
    if (parsedEnd) {
        payload.actualEndTime = parsedEnd;
    }
    if (parsedStart && parsedEnd && new Date(parsedEnd).getTime() < new Date(parsedStart).getTime()) {
        showActualTimeError('End time cannot be before start time.');
        return;
    }
    if (!Object.keys(payload).length) {
        setHeaderEditMode(false);
        return;
    }
    JSUTIL.callAJAXPost(`/data/ProductionOrder/update/${currentBPCRId}`,
        JSON.stringify(payload),
        function () {
            if (currentBPCRRecord) {
                if (payload.name) {
                    currentBPCRRecord.name = payload.name;
                }
                if (payload.actualStartTime) {
                    currentBPCRRecord.actualStartTime = payload.actualStartTime;
                }
                if (payload.actualEndTime) {
                    currentBPCRRecord.actualEndTime = payload.actualEndTime;
                }
            }
            populateBPCRHeader(currentBPCRRecord, currentBPCRId);
            setHeaderEditMode(false);
        },
        function (err) {
            console.error('Failed to save header details', err);
            showActualTimeError('Could not save header details. Please try again.');
        }
    );
}

function saveFinalProductDetails() {
    if (isFinalProductLockedStatus(currentStatusValue)) {
        showFinalProductError('Final product details are locked.');
        return;
    }
    const result = buildFinalProductPayload();
    if (!result.valid) {
        showFinalProductError(result.message || 'Enter valid final product details.');
        return;
    }
    saveFinalProduct(result.payload, function () {
        setFinalProductEditMode(false);
        renderFinalProductSection(finalProductItem);
    }, function () {
        showFinalProductError('Could not save final product details. Please try again.');
    });
}

function buildTestRequestBasePayload() {
    const product = finalProductItem?.product
        || currentBPCRRecord?.productRecipe?.productStage?.product
        || currentBPCRRecord?.productRecipe?.product
        || currentBPCRRecord?.productStage?.product
        || {};
    const productName = product?.name || '';
    const productId = product?.id || product?.Id || '';
    const productType = product?.type
        || product?.Type
        || currentBPCRRecord?.productRecipe?.type
        || '';
    const uom = product?.uom || product?.Uom || '';
    const batchNumber = currentBPCRRecord?.name
        || currentBPCRRecord?.number
        || currentBPCRId
        || '';
    const manufacturingDate = finalProductItem?.manufacturingDate || '';
    const expiryDate = finalProductItem?.expiryDate || '';
    const receivedDate = new Date().toISOString().split('T')[0];
    if (!productName) {
        return { valid: false, message: 'Product name missing for QC request.' };
    }
    if (!productId) {
        return { valid: false, message: 'Product id missing for QC request.' };
    }
    return {
        valid: true,
        payload: {
            batchNumber: batchNumber,
            description: 'Production',
            manufacturingDate: manufacturingDate,
            expiryDate: expiryDate,
            receivedDate: receivedDate,
            requestStatus: 'New',
            sampleCondition: productName,
            sampleId: productId,
            sampleType: productType,
            sampleUnit: uom,
            // productionOrder: { id: Number(currentBPCRId)},
            samplingMethod: currentBPCRId,
            // parent: { id: Number(currentBPCRId)}
        }
    };
}

function buildTestRequestPayload(sampleQuantity, basePayload) {
    const baseResult = basePayload
        ? { valid: true, payload: basePayload }
        : buildTestRequestBasePayload();
    if (!baseResult.valid) {
        return baseResult;
    }
    const trimmedQty = (sampleQuantity || '').toString().trim();
    if (!trimmedQty) {
        return { valid: false, message: 'Enter a sample quantity for QC request.' };
    }
    return {
        valid: true,
        payload: {
            ...baseResult.payload,
            sampleQuantity: trimmedQty
        }
    };
}

function sendToQC() {
    if (!currentBPCRId) {
        return;
    }
    if (!isFinalStageForCurrentBatch()) {
        JSUTIL.buildErrorModal('Quality Control is only available for final stage batches.');
        return;
    }
    const baseResult = buildTestRequestBasePayload();
    if (!baseResult.valid) {
        console.log(baseResult.message || 'Cannot create QC request.');
        return;
    }
    pendingTestRequestBasePayload = baseResult.payload;
    $('#qcSampleQuantityInput').val('0.5');
    $('.qc-sample-error').hide().text('');
    $('#qcSampleQuantityModal').modal('show');
}

function confirmSampleQuantityAndSend() {
    const qty = $('#qcSampleQuantityInput').val();
    const result = buildTestRequestPayload(qty, pendingTestRequestBasePayload);
    if (!result.valid) {
        $('.qc-sample-error').text(result.message || 'Enter a valid sample quantity.').show();
        return;
    }
    $('#qcSampleQuantityModal').modal('hide');
    pendingTestRequestBasePayload = null;
    submitTestRequest(result.payload, qty);
}

function submitTestRequest(payload, sampleQuantityValue) {
    if (!payload) {
        return;
    }
    $('.send-to-qc').prop('disabled', true);
    console.log("test : ", payload)
    const url = '/data/TestRequest/create_with_sequence?field=ReferenceNumber';
    JSUTIL.callAJAXPost(url,
        JSON.stringify(payload),
        function (res) {
            console.log('QC request created.');
            createSampleIssueTransaction(sampleQuantityValue, res);
            markBPCRCompleteThenMoveFinalProduct();
        },
        function (err) {
            console.error('Failed to create TestRequest', err);
            console.log('Could not create QC request. Please try again.');
            $('.send-to-qc').prop('disabled', false);
        });
}

function markBPCRCompleteThenMoveFinalProduct() {
    updateBPCRStatusValue(BPCR_COMPLETED_STATUS, function () {
        updateFinalProductTypeToWarehouse(function () {
            loadBPCRDetails(currentBPCRId);
            $('.send-to-qc').prop('disabled', true);
        }, function () {
            console.log('Could not move final product to Under Test. Please try again.');
            $('.send-to-qc').prop('disabled', false);
        });
    }, function () {
        console.log('Could not mark BPCR as Under Test. Please try again.');
        $('.send-to-qc').prop('disabled', false);
    });
}

function createSampleIssueTransaction(sampleQuantityValue, testRequestResponse) {
    const qty = Math.abs(Number(sampleQuantityValue));
    if (!Number.isFinite(qty) || qty <= 0) {
        return;
    }
    const batchNumber = getBatchNumberForTransactions();
    if (!batchNumber) {
        return;
    }
    const docNumber = testRequestResponse?.referenceNumber
        || testRequestResponse?.ReferenceNumber
        || testRequestResponse?.number
        || testRequestResponse?.Number
        || getBPCRDocNumber()
        || batchNumber;
    const orderMaterialId = finalProductItem?.id || '';
    const payload = {
        docNumber: docNumber,
        batchNumber: batchNumber,
        transactionDate: new Date().toISOString(),
        quantityChange: -qty,
        type: 'Issued'
    };
    if (orderMaterialId) {
        payload.orderMaterial = { id: orderMaterialId };
    }
    JSUTIL.callAJAXPost('/data/OrderMaterialTransaction/create',
        JSON.stringify(payload),
        function () {
            refreshProductTransferList();
        },
        function (err) {
            console.error('Failed to record sample issue transaction', err);
        });
}

function setMaterialEditMode(enable) {
    const allowEdit = canEditMaterials(currentStatusValue);
    materialEditEnabled = enable && allowEdit;
    const canSaveHeaderTimes = canEditActualTimes(currentStatusValue);
    const showSaveButton = materialEditEnabled || canSaveHeaderTimes;
    $('.save-bpcr-material').toggle(showSaveButton);
    $('.edit-bpcr-material').toggle(!materialEditEnabled && allowEdit);
}

function setHeaderEditMode(enable) {
    const normalized = (currentStatusValue || '').toLowerCase();
    const lockedForTest = normalized === 'under test' || normalized === 'under-test';
    headerEditEnabled = !!enable && !lockedForTest;
    $('.header-edit-btn').toggle(!headerEditEnabled);
    $('.header-save-btn').toggle(headerEditEnabled);
    const canEditTimes = canEditActualTimes(currentStatusValue) && headerEditEnabled;
    $('#bpcrNumberDisplay').toggle(!headerEditEnabled);
    $('#bpcrBatchNameInput')
        .toggle(headerEditEnabled)
        .prop('disabled', !headerEditEnabled);
    $('#bpcrActualStartInput, #bpcrActualEndInput').prop('disabled', !canEditTimes);
    toggleActualTimeEditControls(currentStatusValue);
}

function setFinalProductEditMode(enable) {
    const normalized = (currentStatusValue || '').toLowerCase();
    const locked = isFinalProductLockedStatus(currentStatusValue) || normalized === 'under test' || normalized === 'under-test';
    finalProductEditEnabled = !!enable && !locked;
    updateFinalProductButtons(locked, !!$('#finalProductProductId').val());
    toggleFinalProductInputs(finalProductEditEnabled);
}

function updateFinalProductButtons(isLocked, hasProduct) {
    if (isLocked) {
        finalProductEditEnabled = false;
    }
    const canShowEdit = !isLocked && hasProduct;
    const showSave = finalProductEditEnabled && canShowEdit;
    $('.final-product-edit-btn').toggle(canShowEdit && !finalProductEditEnabled);
    $('.final-product-save-btn').toggle(showSave);
}

function toggleFinalProductInputs(canEdit) {
    const locked = isFinalProductLockedStatus(currentStatusValue);
    const allowEdit = canEdit && !locked;
    $('#finalProductQty, #finalProductNumberOfContainers, #finalProductManufacturingDate, #finalProductExpiryDate, #finalProductExpiryType, #finalProductNetQtyPerContainer, #finalProductVariationPerContainer, #finalProductGrade')
        .prop('disabled', !allowEdit)
        .prop('readonly', !allowEdit);
}

function updateFinalProductDateLabel() {
    const selectedType = normalizeExpiryType($('#finalProductExpiryType').val() || 'Expiry');
    const labelText = selectedType === 'Retest' ? 'Retest Date' : 'Expiry Date';
    $('#finalProductExpiryDateLabel').text(labelText);
}

function toggleOwnerVisibility(status) {
    const hideOwner = (status || '').toLowerCase() === 'new';
    // $('#bpcrOwnerWrapper').toggle(!hideOwner);
}

function toggleEditAvailability(status) {
    const allowEdit = canEditMaterials(status);
    if (!allowEdit) {
        setMaterialEditMode(false);
    } else {
        setMaterialEditMode(materialEditEnabled);
    }
}

function canEditMaterials(status) {
    const normalized = (status || '').toLowerCase();
    const isReviewStatus = normalized === 'in review';
    const isRejected = normalized === 'rejected';
    // Rejected: can edit again. Certain statuses are always locked.
    const lockedStatuses = [
        'approved',
        'assigned',
        'rm-request',
        'rm-requested',
        'rm-issued',
        'rm issued',
        'start production',
        'in-production',
        'end production',
        'production ended',
        'product issued',
        'complete',
        'under test',
        'under-test',
        'sent to qa'
    ];
    const isLocked = isReviewStatus || lockedStatuses.indexOf(normalized) !== -1;
    return !isBPCRReviewPage && (isRejected || !isLocked);
}

function canEditRequestedQuantities(stepName) {
    const normalizedStatus = (currentStatusValue || '').toLowerCase();
    const statusAllowsEdit = !normalizedStatus
        || normalizedStatus === 'new'
        || normalizedStatus === 'assigned'
        || normalizedStatus === 'rm-request'
        || normalizedStatus === 'rm-requested'
        || normalizedStatus === 'rm issued'
        || normalizedStatus === 'rm-issued'
        || normalizedStatus === 'product issued';
    if (!statusAllowsEdit) {
        return false;
    }
    const normalizedLot = normalizeLotName(stepName || '');
    const selectedLotInfo = getSelectedLotInfo();
    const selectedLot = normalizeLotName(selectedLotInfo.normalized || selectedLotInfo.label || '');
    const lotToCheck = normalizedLot || selectedLot;
    if (hasExistingMaterialRequest) {
        if (lotToCheck) {
            return !isLotAlreadyRequested(lotToCheck);
        }
        return false;
    }
    return true;
}

function canEditUsedQuantities(statusOverride) {
    const normalized = (statusOverride || currentStatusValue || '').toLowerCase();
    const editableStatuses = [
        'rm issued',
        'rm-issued',
        'product issued',
        'production output',
        'packing list creation',
        'start production',
        'in-production',
        'end production',
        'production ended'
    ];
    const isEditableStage = editableStatuses.indexOf(normalized) !== -1;
    if (!isEditableStage) {
        return false;
    }
    return !isCompletedStatus(normalized);
}

function canEditActualTimes(status) {
    return !isCompletedStatus(status);
}

function canEditPackingList(statusValue) {
    const normalized = (statusValue || currentStatusValue || '').toLowerCase();
    return normalized === 'packing list creation';
}

function isPackingListVisibleStatus(statusValue) {
    return getProgressStepForStatus(statusValue) >= 5;
}

function isInProcessLotSelected() {
    const $filter = $('#bpcrMaterialTypeFilter');
    if (!$filter.length) {
        return false;
    }
    const selectedValue = ($filter.val() || '').toString().trim().toLowerCase();
    const $selectedOption = $filter.find('option:selected');
    const stepName = ($selectedOption.data('step-name') || '').toString().trim().toLowerCase();
    const normalizedValue = getNormalizedMaterialType(selectedValue);
    return normalizedValue === 'inprocess'
        || normalizedValue === 'in-process'
        || stepName === 'in-process';
}

function toggleAssignButton(status) {
    $('.assign-bpcr').hide();
    const normalized = (status || '').toLowerCase();
    const allowRequestByStatus = !normalized
        || normalized === 'new'
        || normalized === 'rm-request'
        || normalized === 'rm-requested'
        || normalized === 'rm issued'
        || normalized === 'rm-issued'
        || normalized === 'product issued';
    const isInProcessLot = isInProcessLotSelected();
    const lotInfo = getSelectedLotInfo();
    const isLotRequested = lotInfo.normalized ? isLotAlreadyRequested(lotInfo.normalized) : false;
    const canRequest = allowRequestByStatus
        && !materialRequestCheckPending
        && !isInProcessLot
        && !isLotRequested;
    $('.request-materials').toggle(canRequest);
    $('.material-request-note').toggle(allowRequestByStatus && isLotRequested && !materialRequestCheckPending);
}

function toggleProductionButtons(status) {
    $('.start-production, .end-production').hide();
    const showUsedQtySave = canEditUsedQuantities(status);
    $('.save-used-qty').toggle(showUsedQtySave);
}

function toggleFinalProductTab(status) {
    const normalized = (status || '').toLowerCase();
    const showTab = normalized === 'rm-request'
        || normalized === 'rm-requested'
        || normalized === 'rm issued'
        || normalized === 'rm-issued'
        || normalized === 'production output'
        || normalized === 'packing list creation'
        || normalized === 'product issued'
        || isProductionEndedStatus(normalized)
        || isCompletedStatus(normalized);
    const $tab = $('#bpcrFinalProductTab');
    const $content = $('#bpcr-final-product');
    const $materialTab = $('#bpcrMaterialTab');
    const $materialContent = $('#bpcr-material');
    if (!$tab.length || !$content.length) {
        return;
    }
    if (showTab) {
        $tab.show();
    } else {
        const isActive = $tab.hasClass('active') || $content.hasClass('active');
        $tab.removeClass('active');
        $content.removeClass('active');
        $tab.hide();
        if (isActive) {
            $materialContent.addClass('active');
            $materialTab.addClass('active');
        }
    }
}

function toggleQCButton(status) {
    const isFinalStage = isFinalStageForCurrentBatch();
    const normalized = (status || '').toLowerCase();
    const isRelease = isReleaseStatus(normalized);
    const showQC = isFinalStage && !isRelease && (normalized === 'production output'
        || normalized === 'packing list creation'
        || isProductionEndedStatus(normalized)
        || isCompletedStatus(normalized));
    const hideForTest = normalized === 'under test' || normalized === 'under-test';
    $('.send-to-qc').toggle(showQC && !hideForTest);
    const showReport = showQC
        || isCompletedStatus(normalized)
        || (!isFinalStage && (normalized === 'production output' || isProductionEndedStatus(normalized)));
    $('.print-bpcr-report').toggle(showReport);
    $('.print-final-label').toggle(showQC);
    toggleFinishProductionButton(status);
}

function toggleActualTimeEditControls(status) {
    const statusAllows = canEditActualTimes(status);
    const canEdit = statusAllows && headerEditEnabled;
    $('#bpcrActualStartInput, #bpcrActualEndInput').prop('disabled', !canEdit);
    $('.actual-time-lock-text').toggle(!statusAllows);
}

function toggleTabsVisibility(status) {
    const $tabs = $('#bpcrTabs');
    if (!$tabs.length) {
        return;
    }
    const hideTabs = (status || '').toLowerCase() === 'sent to qa';
    $tabs.toggle(!hideTabs);
}

function toggleProductTransferButtons(status) {
    const normalized = (status || '').toLowerCase();
    const progressStep = getProgressStepForStatus(status);
    const isFinalStage = isFinalStageForCurrentBatch();
    const isNonFinalStage = !isFinalStage;
    const showIntermediate = normalized === 'new';
    const showAddTransfer = progressStep >= 6
        || normalized === 'qc-approved'
        || (isNonFinalStage && progressStep >= 5);
    $('.get-intermediate-btn').toggle(showIntermediate);
    $('.add-transfer-btn').toggle(showAddTransfer && !isFinalStage);
}

function getProgressStepForStatus(status) {
    const normalized = (status || '').toLowerCase();
    if (!normalized) {
        return 1;
    }
    if (normalized === 'new' || normalized === 'assigned') {
        return 3;
    }
    if (normalized === 'sent to qa' || normalized === 'in review' || normalized === 'qa review') {
        return 2;
    }
    if (normalized === 'rm-request' || normalized === 'rm-requested') {
        return 3;
    }
    if (normalized === 'rm issued' || normalized === 'rm-issued' || normalized === 'product issued' || normalized === 'start production' || normalized === 'in-production' || normalized === 'production output') {
        return 4;
    }
    if (normalized === 'end production' || normalized === 'production ended') {
        return 4;
    }
    if (normalized === 'packing list creation') {
        return 5;
    }
    if (normalized === 'under test' || normalized === 'under-test' || normalized === 'qc-approved' || normalized === 'qc-rejected') {
        return 6;
    }
    if (normalized === 'complete' || normalized === 'release' || normalized === 'released') {
        return 7;
    }
    return 1;
}

function updateProgressSteps(status) {
    const currentStep = getProgressStepForStatus(status);
    const $steps = $('.step-button');
    $steps.each(function () {
        var stepNumber = Number($(this).data('step'));
        $(this).removeClass('active completed');
        const isCompleted = stepNumber < currentStep;
        const isActive = stepNumber === currentStep;
        $(this).toggleClass('completed', isCompleted);
        $(this).toggleClass('active', isActive);
    });
}

function getBatchNumberForTransactions() {
    return (currentBPCRRecord?.name
        || currentBPCRRecord?.number
        || $('#bpcrBatchNameInput').val()
        || '').toString().trim();
}

function refreshProductTransferList() {
    const batchNumber = getBatchNumberForTransactions();
    if (!batchNumber) {
        renderProductTransferTable([]);
        return;
    }
    fetchProductTransferTransactions(batchNumber);
}

function openAddTransferModal() {
    if (!currentBPCRId) {
        return;
    }
    const batchNumber = getBatchNumberForTransactions();
    if (!batchNumber) {
        JSUTIL.buildErrorModal('Batch number missing. Please refresh and try again.');
        return;
    }
    const defaultDoc = ($('#finalProductDocNumber').val() || getBPCRDocNumber() || batchNumber || '').toString().trim();
    $('#bpcrTransferDocCode').val(defaultDoc || batchNumber);
    $('#bpcrTransferDate').val(formatDateInputValue(new Date().toISOString()));
    $('#bpcrTransferQuantity').val('');
    $('#bpcrTransferTypeSelect').val('QC');
    $('#bpcrTransferTargetBatchName').val('');
    updateTransferTypeOptions();
    loadTransferPlantOptions();
    setTransferModalMessage('', false);
    updateAvailableTransferQuantityDisplay();
    $('#bpcrAddTransferModal').modal('show');
    
}

function resetTransferModal() {
    $('#bpcrTransferDocCode').val('');
    $('#bpcrTransferDate').val('');
    $('#bpcrTransferQuantity').val('');
    $('#bpcrTransferTypeSelect').val('QC');
    $('#bpcrTransferPlantSelect').val('');
    $('#bpcrTransferTargetBatchName').val('');
    $('#bpcrTransferAvailableQty').val('');
    $('.bpcr-transfer-plant-group').hide();
    $('.bpcr-transfer-target-group').hide();
    setTransferModalMessage('', false);
}

function updateTransferTypeOptions() {
    const $select = $('#bpcrTransferTypeSelect');
    if (!$select.length) {
        return;
    }
    const hasBsrOption = isFinalStageForCurrentBatch();
    const currentValue = ($select.val() || '').toLowerCase();
    const $bsrOption = $select.find('option[value=\"BSR\"]');
    if (hasBsrOption && !$bsrOption.length) {
        $select.append('<option value=\"BSR\">BSR</option>');
    } else if (!hasBsrOption && $bsrOption.length) {
        $bsrOption.remove();
    }
    const availableValues = $select.find('option').map(function () {
        return ($(this).val() || '').toLowerCase();
    }).get();
    const preferredValue = availableValues.indexOf(currentValue) !== -1
        ? $select.val()
        : 'QC';
    $select.val(preferredValue);
    $select.trigger('change');
}

function handleTransferTypeChange() {
    const typeValue = ($('#bpcrTransferTypeSelect').val() || '').toLowerCase();
    const isPlant = typeValue === 'plant';
    $('.bpcr-transfer-plant-group').toggle(isPlant);
    if (!isPlant) {
        $('#bpcrTransferPlantSelect').val('');
        $('#bpcrTransferTargetBatchName').val('');
        $('.bpcr-transfer-target-group').hide();
    }
    handleTransferPlantChange();
}

function handleTransferPlantChange() {
    const typeValue = ($('#bpcrTransferTypeSelect').val() || '').toLowerCase();
    if (typeValue !== 'plant') {
        $('.bpcr-transfer-target-group').hide();
        return;
    }
    const plantId = $('#bpcrTransferPlantSelect').val();
    if (!plantId) {
        $('.bpcr-transfer-target-group').hide();
        return;
    }
    $('.bpcr-transfer-target-group').show();
}

function loadTransferPlantOptions() {
    const payload = {
        fields: 'Id;Name;Number'
    };
    JSUTIL.callAJAXPost('/data/Plant/list',
        JSON.stringify(payload),
        function (res) {
            const list = Array.isArray(res?.data) ? res.data : (Array.isArray(res) ? res : []);
            transferPlantOptions = list.map(function (item) {
                const id = item.id || item.Id || '';
                if (!id) {
                    return null;
                }
                const code = item.code || item.Code || '';
                const number = item.number || item.Number || '';
                const name = item.name || item.Name || '';
                const parts = [];
                if (code) {
                    parts.push(code);
                } else if (number) {
                    parts.push(number);
                }
                if (name) {
                    parts.push(name);
                }
                return {
                    id: String(id),
                    label: parts.join(' - ') || name || ('Plant ' + id)
                };
            }).filter(Boolean);
            renderTransferPlantOptions();
        },
        function (err) {
            console.error('Failed to load plants for transfer', err);
            transferPlantOptions = [];
            renderTransferPlantOptions();
        }
    );
}

function renderTransferPlantOptions() {
    const $select = $('#bpcrTransferPlantSelect');
    if (!$select.length) {
        return;
    }
    const options = ['<option value="">Select plant</option>'].concat(
        transferPlantOptions.map(function (opt) {
            return `<option value="${opt.id}">${opt.label}</option>`;
        })
    );
    $select.html(options.join(''));
    $('.select2').select2();
}

function getAvailableTransferQuantity() {
    const transactions = Array.isArray(productTransferTransactions)
        ? productTransferTransactions
        : [];
    if (!transactions.length) {
        const fallbackQty = firstDefinedNumber(
            finalProductItem?.plannedQuantity,
            $('#finalProductQty').val(),
            0
        );
        return Number.isFinite(fallbackQty) ? Math.max(Number(fallbackQty), 0) : 0;
    }
    const balance = transactions.reduce(function (running, txn) {
        const quantities = getTransferQuantities(txn);
        return running + (quantities.receiptValue || 0) - (quantities.issueValue || 0);
    }, 0);
    return Math.max(balance, 0);
}

function updateAvailableTransferQuantityDisplay() {
    const $input = $('#bpcrTransferAvailableQty');
    if (!$input.length) {
        return;
    }
    const availableQty = getAvailableTransferQuantity();
    const uom = finalProductItem?.product?.uom || finalProductItem?.product?.Uom || '';
    const formattedQty = formatQtyValue(availableQty);
    const displayValue = uom
        ? `${formattedQty || '0'} ${uom}`
        : (formattedQty || '0');
    $input.val(displayValue);
}

function saveTransferTransaction() {
    const qtyValue = Number($('#bpcrTransferQuantity').val());
    if (!Number.isFinite(qtyValue) || qtyValue <= 0) {
        setTransferModalMessage('Enter a valid quantity.', true);
        return;
    }
    const availableQty = getAvailableTransferQuantity();
    const uomLabel = finalProductItem?.product?.uom || finalProductItem?.product?.Uom || '';
    if (!Number.isFinite(availableQty) || availableQty <= 0) {
        setTransferModalMessage('No quantity available to transfer.', true);
        return;
    }
    if (qtyValue > availableQty) {
        const availableDisplay = formatQtyValue(availableQty) || '0';
        const suffix = uomLabel ? ` ${uomLabel}` : '';
        setTransferModalMessage(`Quantity exceeds available (${availableDisplay}${suffix}).`, true);
        return;
    }
    const typeValue = ($('#bpcrTransferTypeSelect').val() || '').toLowerCase();
    const isPlantTransfer = typeValue === 'plant';
    const isBsrTransfer = typeValue === 'bsr';
    if (isBsrTransfer && !isFinalStageForCurrentBatch()) {
        setTransferModalMessage('BSR transfer is only available for final stage batches.', true);
        return;
    }
    let transferLabel = '';
    let targetBatchName = '';
    let targetPlantId = '';
    if (isPlantTransfer) {
        const plantId = $('#bpcrTransferPlantSelect').val();
        targetPlantId = plantId;
        const targetName = ($('#bpcrTransferTargetBatchName').val() || '').toString().trim();
        if (!plantId) {
            setTransferModalMessage('Select a plant.', true);
            return;
        }
        if (!targetName) {
            setTransferModalMessage('Enter target batch name.', true);
            return;
        }
        const plantLabel = getTransferOptionLabel(transferPlantOptions, plantId);
        transferLabel = `${plantLabel || 'Plant'} | ${targetName}`;
        targetBatchName = targetName;
    } else if (isBsrTransfer) {
        transferLabel = 'BSR';
    } else {
        transferLabel = 'QC';
    }
    const batchNumber = getBatchNumberForTransactions();
    const docCodeInput = ($('#bpcrTransferDocCode').val() || '').toString().trim();
    const docCode = docCodeInput || batchNumber || getBPCRDocNumber() || '';
    const dateValue = $('#bpcrTransferDate').val();
    const transactionDate = dateValue ? new Date(dateValue).toISOString() : new Date().toISOString();
    const orderMaterialId = finalProductItem?.id || '';
    const buildTxnPayload = function (changeReasonId) {
        const payload = {
            docNumber: docCode,
            transactionDate: new Date().toISOString(),
            batchNumber: batchNumber,
            quantityChange: -Math.abs(qtyValue),
            type: 'Issued',
            issuedToName: transferLabel
        };
        if (orderMaterialId) {
            payload.orderMaterial = { id: orderMaterialId };
        }
        if (targetBatchName) {
            payload.targetBatchName = targetBatchName;
        }
        if (targetPlantId) {
            payload.targetPlant = { id: targetPlantId };
        }
        if (changeReasonId) {
            payload.changeReason = changeReasonId;
        }
        return payload;
    };
    const submitTransaction = function (changeReasonId) {
        const payload = buildTxnPayload(changeReasonId);
        $('#bpcrTransferSaveBtn').prop('disabled', true);
        JSUTIL.callAJAXPost('/data/OrderMaterialTransaction/create',
            JSON.stringify(payload),
            function (res) {
                $('#bpcrAddTransferModal').modal('hide');
                refreshProductTransferList();
                $('#bpcrTransferSaveBtn').prop('disabled', false);
            },
            function (err) {
                console.error('Failed to create product transfer transaction', err);
                setTransferModalMessage('Failed to save transfer. Please try again.', true);
                if (changeReasonId) {
                    deleteTransferOrderMaterial(changeReasonId, function () { }, function () { });
                }
                $('#bpcrTransferSaveBtn').prop('disabled', false);
            }
        );
    };
    if (isBsrTransfer) {
        const bsrResult = buildBsrOrderMaterialPayload(qtyValue, docCode);
        if (!bsrResult.valid) {
            setTransferModalMessage(bsrResult.message || 'Enter valid details for BSR transfer.', true);
            return;
        }
        $('#bpcrTransferSaveBtn').prop('disabled', true);
        createBsrOrderMaterial(bsrResult.payload,
            function (createdId) {
                submitTransaction(createdId);
            },
            function () {
                setTransferModalMessage('Failed to create BSR stock entry. Please try again.', true);
                $('#bpcrTransferSaveBtn').prop('disabled', false);
            }
        );
        return;
    }
    submitTransaction();
}

function buildBsrOrderMaterialPayload(issueQty, docNumberValue) {
    const qty = Math.abs(Number(issueQty));
    if (!Number.isFinite(qty) || qty <= 0) {
        return { valid: false, message: 'Enter a valid quantity.' };
    }
    const baseItem = finalProductItem || {};
    const productId = baseItem?.product?.id
        || baseItem?.product?.Id
        || $('#finalProductProductId').val()
        || '';
    if (!productId) {
        return { valid: false, message: 'Final product is missing. Save final product details first.' };
    }
    const docNumber = (docNumberValue
        || baseItem.docNumber
        || $('#finalProductDocNumber').val()
        || getBPCRDocNumber()
        || '').toString().trim();
    if (!docNumber) {
        return { valid: false, message: 'Doc number missing for BSR entry.' };
    }
    const manufacturingDate = baseItem.manufacturingDate
        || $('#finalProductManufacturingDate').val()
        || '';
    const expiryDate = baseItem.expiryDate
        || $('#finalProductExpiryDate').val()
        || '';
    const retestDate = baseItem.retestDate
        || baseItem.expiryDate
        || $('#finalProductExpiryDate').val()
        || '';
    const expiryType = normalizeExpiryType(
        baseItem.expiryType
        || $('#finalProductExpiryType').val()
        || 'Expiry'
    );
    const productCodeValue = buildWarehouseProductCodeValue(baseItem);
    const payload = {
        type: BSR_ORDER_MATERIAL_TYPE,
        plannedQuantity: qty,
        quantityToWarehouse: qty,
        materialStatus: baseItem.materialStatus || 'Actual',
        remarks: baseItem.remarks || currentBPCRId,
        productLot: BSR_PRODUCT_LOT,
        product: { id: Number(productId) },
        docNumber: docNumber
    };
    if (productCodeValue) {
        payload.productCode = productCodeValue;
    }
    if (manufacturingDate) {
        payload.manufacturingDate = manufacturingDate;
    }
    if (expiryDate) {
        payload.expiryDate = expiryDate;
    }
    if (retestDate) {
        payload.retestDate = retestDate;
    }
    payload.expiryType = expiryType;
    const containerCount = firstDefinedNumber(
        baseItem.numberOfContainers,
        $('#finalProductNumberOfContainers').val(),
        0
    );
    if (containerCount > 0) {
        payload.numberOfContainers = containerCount;
    }
    return { valid: true, payload: payload };
}

function createBsrOrderMaterial(payload, onSuccess, onError) {
    JSUTIL.callAJAXPost('/data/OrderMaterial/create',
        JSON.stringify(payload),
        function (res) {
            const createdId = res?.id || res?.Id || res?.data?.id || res?.data?.Id || '';
            if (typeof onSuccess === 'function') {
                onSuccess(createdId, res);
            }
        },
        function (err) {
            console.error('Failed to create BSR order material', err);
            if (typeof onError === 'function') {
                onError(err);
            }
        }
    );
}

function getTransferOptionLabel(options, id) {
    const match = (Array.isArray(options) ? options : []).find(function (opt) {
        return String(opt.id) === String(id);
    });
    return match ? match.label : '';
}

function setTransferModalMessage(message, isError) {
    const $msg = $('#bpcrTransferMessage');
    if (!$msg.length) {
        return;
    }
    if (!message) {
        $msg.text('').hide();
        return;
    }
    $msg.removeClass('text-danger text-success')
        .addClass(isError ? 'text-danger' : 'text-success')
        .text(message)
        .show();
}

function fetchProductTransferTransactions(batchNumber) {
    const payload = {
        fields: [
            'Id',
            'DocNumber',
            'BatchNumber',
            'TransactionDate',
            'QuantityChange',
            'Type',
            'Arn',
            'IssuedTo.Name',
            'ChangeReason',
            'TargetBatch.Id',
            'TargetBatchName',
            'TargetPlant.Id',
            'OrderMaterial.Id',
            'IssuedToName'
        ].join(';'),
        conditions: [`BatchNumber = '${batchNumber}'`],
        logic: '{0}'
    };
    JSUTIL.callAJAXPost('/data/OrderMaterialTransaction/list',
        JSON.stringify(payload),
        function (res) {
            const list = Array.isArray(res?.data) ? res.data : (Array.isArray(res) ? res : []);
            productTransferTransactions = (list || []).slice().sort(function (a, b) {
                const aTime = new Date(a.transactionDate || a.TransactionDate || 0).getTime();
                const bTime = new Date(b.transactionDate || b.TransactionDate || 0).getTime();
                return aTime - bTime;
            });
            renderProductTransferTable(productTransferTransactions);
            updateAvailableTransferQuantityDisplay();
        },
        function (err) {
            console.error('Failed to fetch product transfer transactions', err);
            productTransferTransactions = [];
            renderProductTransferTable([]);
            updateAvailableTransferQuantityDisplay();
        }
    );
}

function openIntermediateModal() {
    if (!currentBPCRId) {
        return;
    }
    const batchNumber = getBatchNumberForTransactions();
    if (!batchNumber) {
        JSUTIL.buildErrorModal('Batch number missing. Please refresh and try again.');
        return;
    }
    intermediateTransferTransactions = [];
    setIntermediateModalMessage('', false);
    renderIntermediateTransferTable([], true);
    $('#bpcrIntermediateModal').modal('show');
    fetchIntermediateTransferTransactions(batchNumber);
}

function fetchIntermediateTransferTransactions(batchNumber) {
    const targetPlantId = getBPCRPlantId(currentBPCRRecord);
    const payload = {
        fields: [
            'Id',
            'DocNumber',
            'BatchNumber',
            'TransactionDate',
            'QuantityChange',
            'Type',
            'OrderMaterial.Id',
            'TargetBatchName',
            'TargetPlant.Id',
            'IssuedToName',
            'ChangeReason'
        ].join(';'),
        conditions: ["Type = 'Issued'"],
        logic: '{0}'
    };
    renderIntermediateTransferTable([], true);
    JSUTIL.callAJAXPost('/data/OrderMaterialTransaction/list',
        JSON.stringify(payload),
        function (res) {
            const list = Array.isArray(res?.data) ? res.data : (Array.isArray(res) ? res : []);
            intermediateList = list;
            const filtered = filterIntermediateTransfers(list, batchNumber, targetPlantId);
            intermediateTransferTransactions = filtered;
            renderIntermediateTransferTable(filtered, false);
            if (!filtered.length) {
                setIntermediateModalMessage('No transfers found for this batch/plant.', false);
            } else {
                setIntermediateModalMessage('', false);
            }
        },
        function (err) {
            console.error('Failed to fetch intermediate transfers', err);
            intermediateTransferTransactions = [];
            renderIntermediateTransferTable([], false);
            setIntermediateModalMessage('Failed to load transfers. Please try again.', true);
        }
    );
}

function filterIntermediateTransfers(list, batchNumber, plantId) {
    const normalizedBatchNumber = (batchNumber || '').toString().trim().toLowerCase();
    const normalizedBatchId = (currentBPCRRecord?.id
        || currentBPCRRecord?.Id
        || currentBPCRId
        || '').toString().trim().toLowerCase();
    const normalizedPlantId = plantId ? String(plantId).trim().toLowerCase() : '';
    const normalizedPlantNames = new Set();
    const plantName = getBPCRPlantName(currentBPCRRecord);
    if (plantName) {
        normalizedPlantNames.add(plantName.toString().trim().toLowerCase());
    }
    const plantNumber = currentBPCRRecord?.plant?.number
        || currentBPCRRecord?.plant?.Number
        || currentBPCRRecord?.productRecipe?.plant?.number
        || currentBPCRRecord?.productRecipe?.plant?.Number
        || currentBPCRRecord?.productRecipe?.productStage?.plant?.number
        || currentBPCRRecord?.productRecipe?.productStage?.plant?.Number
        || '';
    if (plantNumber) {
        normalizedPlantNames.add(plantNumber.toString().trim().toLowerCase());
    }
    const plantCode = currentBPCRRecord?.plant?.code
        || currentBPCRRecord?.plant?.Code
        || '';
    if (plantCode) {
        normalizedPlantNames.add(plantCode.toString().trim().toLowerCase());
    }
    return (Array.isArray(list) ? list : []).filter(function (txn) {
        const typeValue = (txn.type || txn.Type || '').toString().trim().toLowerCase();
        const quantities = getTransferQuantities(txn);
        const qtyValue = quantities.issueValue || quantities.receiptValue || 0;
        const targetPlantId = txn.targetPlant?.id
            || txn.targetPlant?.Id
            || txn.TargetPlant?.Id
            || txn.TargetPlant?.id
            || '';
        const targetPlantName = txn.targetPlant?.name
            || txn.targetPlant?.Name
            || txn.TargetPlant?.Name
            || '';
        const targetPlantNumber = txn.targetPlant?.number
            || txn.targetPlant?.Number
            || txn.TargetPlant?.Number
            || '';
        const targetPlantCode = txn.targetPlant?.code
            || txn.targetPlant?.Code
            || '';
        const targetBatchId = txn.targetBatch?.id
            || txn.targetBatch?.Id
            || txn.TargetBatch?.id
            || txn.TargetBatch?.Id
            || '';
        const targetBatchName = txn.targetBatchName
            || txn.TargetBatchName
            || '';
        const hasChangeReason = !!(txn.changeReason || txn.ChangeReason);
        const normalizedTargetBatchId = targetBatchId.toString().trim().toLowerCase();
        const normalizedTargetBatchName = targetBatchName.toString().trim().toLowerCase();
        const plantMatches = (normalizedPlantId && String(targetPlantId).trim().toLowerCase() === normalizedPlantId)
            || (targetPlantName && normalizedPlantNames.has(targetPlantName.toString().trim().toLowerCase()))
            || (targetPlantNumber && normalizedPlantNames.has(targetPlantNumber.toString().trim().toLowerCase()))
            || (targetPlantCode && normalizedPlantNames.has(targetPlantCode.toString().trim().toLowerCase()));
        const batchMatches = (!!normalizedBatchNumber && normalizedTargetBatchName === normalizedBatchNumber)
            || (!!normalizedBatchId && normalizedTargetBatchId === normalizedBatchId);
        return typeValue === 'issued' && qtyValue > 0 && plantMatches && batchMatches && !hasChangeReason;
    }).sort(function (a, b) {
        const aTime = new Date(a.transactionDate || a.TransactionDate || 0).getTime();
        const bTime = new Date(b.transactionDate || b.TransactionDate || 0).getTime();
        return aTime - bTime;
    });
}

function renderIntermediateTransferTable(list, isLoading) {
    const $tbody = $('#bpcrIntermediateTableBody');
    if (!$tbody.length) {
        return;
    }
    if (isLoading) {
        $tbody.html('<tr><td colspan="7" class="text-center text-muted">Loading transfers...</td></tr>');
        return;
    }
    const rows = (Array.isArray(list) ? list : []).map(function (txn) {
        const txnId = txn.id || txn.Id || '';
        const quantities = getTransferQuantities(txn);
        const qtyValue = quantities.issueValue || quantities.receiptValue;
        if (!Number.isFinite(qtyValue) || qtyValue <= 0) {
            return '';
        }
        const hasChangeReason = !!(txn.changeReason || txn.ChangeReason);
        if (hasChangeReason) {
            return '';
        }
        const qtyDisplay = formatQtyValue(qtyValue);
        const productInfo = getTransferProductInfo(txn);
        const productLabel = [productInfo.name, productInfo.number].filter(Boolean).join(' | ') || '-';
        const batchFrom = txn.batchNumber || txn.BatchNumber || '';
        const docCode = txn.docNumber || txn.DocNumber || batchFrom;
        const targetBatch = txn.targetBatchName || txn.TargetBatchName || '';
        const targetPlant = txn.targetPlant?.name
            || txn.targetPlant?.Name
            || txn.TargetPlant?.Name
            || '';
        const alreadyAdded = txn._addedToTarget
            || intermediateAddedTxnIds.has(String(txnId));
        if (alreadyAdded) {
            return '';
        }
        const qtyWithUom = productInfo.uom ? `${qtyDisplay} ${productInfo.uom}` : qtyDisplay;
        const safeTxnId = escapeHtml(txnId);
        const docDisplay = escapeHtml(docCode || '-');
        const batchDisplay = escapeHtml(batchFrom || '-');
        const productDisplay = escapeHtml(productLabel || '-');
        const qtyDisplayHtml = escapeHtml(qtyWithUom || '-');
        const targetDisplay = escapeHtml(targetPlant || targetBatch || '-');
        const actionCell = `<button type="button" class="btn btn-success btn-xs bpcr-intermediate-get" data-txn-id="${safeTxnId}">Get</button>`;
        return `
            <tr data-txn-id="${safeTxnId}">
                <td>${formatDateShortDisplay(txn.transactionDate || txn.TransactionDate)}</td>
                <td>${docDisplay}</td>
                <td>${batchDisplay}</td>
                <td>${productDisplay}</td>
                <td class="text-right">${qtyDisplayHtml}</td>
                <td>${targetDisplay}</td>
                <td>${actionCell}</td>
            </tr>
        `;
    }).filter(Boolean);
    if (!rows.length) {
        $tbody.html('<tr><td colspan="7" class="text-center text-muted">No transfers found for this batch/plant.</td></tr>');
        return;
    }
    $tbody.html(rows.join(''));
}

function setIntermediateModalMessage(message, isError) {
    const $msg = $('#bpcrIntermediateMessage');
    if (!$msg.length) {
        return;
    }
    if (!message) {
        $msg.text('').hide();
        return;
    }
    $msg.removeClass('text-danger text-success')
        .addClass(isError ? 'text-danger' : 'text-success')
        .text(message)
        .show();
}

function handleIntermediateGet(txnId) {
    if (!txnId) {
        return;
    }
    const txn = (Array.isArray(intermediateTransferTransactions) ? intermediateTransferTransactions : [])
        .find(function (item) { return String(item.id || item.Id) === String(txnId); });
    if (!txn) {
        return;
    }
    const alreadyAdded = txn._addedToTarget
        || intermediateAddedTxnIds.has(String(txnId));
    if (alreadyAdded) {
        setIntermediateModalMessage('This transfer is already added as an in-process material.', false);
        renderIntermediateTransferTable(intermediateTransferTransactions, false);
        return;
    }
    const quantities = getTransferQuantities(txn);
    const qtyValue = quantities.issueValue || quantities.receiptValue;
    if (!Number.isFinite(qtyValue) || qtyValue <= 0) {
        setIntermediateModalMessage('Quantity not available for this transfer.', true);
        return;
    }
    const productInfo = getTransferProductInfo(txn);
    const productId = productInfo.id
        || txn?.product?.id
        || txn?.product?.Id
        || txn?.orderMaterial?.product?.id
        || txn?.orderMaterial?.product?.Id
        || txn?.OrderMaterial?.Product?.Id
        || '';
    if (!productId) {
        setIntermediateModalMessage('Product not found for this transfer.', true);
        return;
    }
    const remarksValue = currentBPCRId || getBatchNumberForTransactions() || '';
    const productLotValue = 'In-Process';
    const payload = {
        type: 'In-Process',
        plannedQuantity: Math.abs(Number(qtyValue)),
        remarks: remarksValue,
        product: { id: productId }
    };
    payload.productLot = productLotValue;
    const $btn = $(`.bpcr-intermediate-get[data-txn-id="${txnId}"]`);
    $btn.prop('disabled', true);
    setIntermediateModalMessage('', false);
    JSUTIL.callAJAXPost('/data/OrderMaterial/create',
        JSON.stringify(payload),
        function (res) {
            const createdItem = normalizeMaterialItem(res?.data || res || {});
            const mergedItem = {
                ...createdItem,
                type: createdItem.type || 'In-Process',
                plannedQuantity: createdItem.plannedQuantity || Math.abs(Number(qtyValue)),
                product: createdItem.product?.id ? createdItem.product : {
                    id: productId,
                    name: productInfo.name,
                    number: productInfo.number,
                    uom: productInfo.uom
                },
                remarks: createdItem.remarks || remarksValue,
                productLot: createdItem.productLot || productLotValue
            };
            bpcrMaterialItems = Array.isArray(bpcrMaterialItems)
                ? bpcrMaterialItems.concat([mergedItem])
                : [mergedItem];
            txn._addedToTarget = true;
            if (txnId) {
                intermediateAddedTxnIds.add(String(txnId));
            }
            if (txnId && mergedItem.id) {
                markIntermediateTxnUsed(txnId, mergedItem.id);
                txn.changeReason = mergedItem.id;
            }
            renderBPCRMaterialTables(bpcrMaterialItems);
            renderIntermediateTransferTable(intermediateTransferTransactions, false);
            setIntermediateModalMessage('In-process material added from transfer.', false);
            applyMaterialTypeFilter();
        },
        function (err) {
            console.error('Failed to create in-process material from transfer', err);
            setIntermediateModalMessage('Failed to create in-process material. Please try again.', true);
            $btn.prop('disabled', false);
        }
    );
}

function markIntermediateTxnUsed(txnId, orderMaterialId) {
    if (!txnId || !orderMaterialId) {
        return;
    }
    const payload = {
        changeReason: String(orderMaterialId)
    };
    JSUTIL.callAJAXPost(`/data/OrderMaterialTransaction/update/${txnId}`,
        JSON.stringify(payload),
        function () { },
        function (err) {
            console.error('Failed to update intermediate transaction changeReason', err);
        }
    );
}

function handleRecallTransfer($button) {
    if (!$button || !$button.length) {
        return;
    }
    const qtyValue = Number($button.data('txnQty'));
    const transferLabel = ($button.data('transferTo') || '').toString();
    if (!Number.isFinite(qtyValue) || qtyValue <= 0) {
        return;
    }
    const batchNumber = getBatchNumberForTransactions();
    if (!batchNumber) {
        JSUTIL.buildErrorModal('Batch number missing. Please refresh and try again.');
        return;
    }
    const confirmMsg = `Recall ${qtyValue} back to batch ${batchNumber}?`;
    if (!window.confirm(confirmMsg)) {
        return;
    }
    const orderMaterialId = finalProductItem?.id || '';
    const payload = {
        docNumber: batchNumber,
        transactionDate: new Date().toISOString(),
        quantityChange: Math.abs(qtyValue),
        type: 'Received',
        batchNumber: `Recall from ${transferLabel || 'batch'}`
    };
    if (orderMaterialId) {
        payload.orderMaterial = { id: orderMaterialId };
    }
    const toggleDisabled = function (state) {
        $button.prop('disabled', !!state);
    };
    toggleDisabled(true);
    JSUTIL.callAJAXPost('/data/OrderMaterialTransaction/create',
        JSON.stringify(payload),
        function () {
            refreshProductTransferList();
            const txnId = $button.data('txnId') || '';
            deleteTargetBatchOrderMaterialForTxn(txnId);
            toggleDisabled(false);
        },
        function (err) {
            console.error('Failed to create recall transaction', err);
            JSUTIL.buildErrorModal('Failed to recall transfer. Please try again.');
            toggleDisabled(false);
        }
    );
}

function handleDeleteTransfer($button) {
    if (!$button || !$button.length) {
        return;
    }
    const txnId = $button.data('txnId');
    if (!txnId) {
        return;
    }
    const changeReasonId = $button.data('changeReason') || '';
    const docCode = ($button.data('docCode') || '').toString().trim();
    const confirmMsg = docCode
        ? `Delete transfer ${docCode}?`
        : 'Delete this transfer transaction?';
    if (!window.confirm(confirmMsg)) {
        return;
    }
    const toggleDisabled = function (state) {
        $button.prop('disabled', !!state);
    };
    toggleDisabled(true);
    const deleteTxn = function () {
        deleteTransferTransaction(txnId,
            function () {
                refreshProductTransferList();
                toggleDisabled(false);
            },
            function () {
                JSUTIL.buildErrorModal('Failed to delete transfer transaction. Please try again.');
                toggleDisabled(false);
            }
        );
    };
    deleteTransferOrderMaterial(changeReasonId, deleteTxn, function () {
        JSUTIL.buildErrorModal('Failed to delete linked order material. Transaction not deleted.');
        toggleDisabled(false);
    });
}

function deleteTransferOrderMaterial(orderMaterialId, onSuccess, onError) {
    if (!orderMaterialId) {
        if (typeof onSuccess === 'function') {
            onSuccess();
        }
        return;
    }
    JSUTIL.callAJAXPost(`/data/OrderMaterial/delete?id=${orderMaterialId}`,
        '{}',
        function () {
            if (typeof onSuccess === 'function') {
                onSuccess();
            }
        },
        function (err) {
            console.error('Failed to delete order material for transfer', err);
            if (typeof onError === 'function') {
                onError(err);
            }
        }
    );
}

function deleteTransferTransaction(txnId, onSuccess, onError) {
    if (!txnId) {
        if (typeof onError === 'function') {
            onError();
        }
        return;
    }
    JSUTIL.callAJAXPost(`/data/OrderMaterialTransaction/delete?id=${txnId}`,
        '{}',
        function () {
            if (typeof onSuccess === 'function') {
                onSuccess();
            }
        },
        function (err) {
            console.error('Failed to delete transfer transaction', err);
            if (typeof onError === 'function') {
                onError(err);
            }
        }
    );
}

function buildTransferTxnRemark(txnId) {
    return txnId ? `${txnId}` : '';
}

function createTargetBatchOrderMaterial(targetBatchId, qtyValue, txnId) {
    const qty = Math.abs(Number(qtyValue));
    const productId = finalProductItem?.product?.id || finalProductItem?.product?.Id || '';
    if (!qty || !targetBatchId || !productId) {
        return;
    }
    const payload = {
        type: 'In-Process',
        plannedQuantity: qty,
        remarks: String(targetBatchId || ''),
        product: { id: productId }
    };
    JSUTIL.callAJAXPost('/data/OrderMaterial/create',
        JSON.stringify(payload),
        function () { },
        function (err) {
            console.error('Failed to create target batch order material', err);
        }
    );
}

function deleteTargetBatchOrderMaterialForTxn(txnId) {
    const remark = buildTransferTxnRemark(txnId);
    const payload = {
        fields: 'Id',
        conditions: remark
            ? [`Remarks = '${remark}'`, "Type = 'In-Process'"]
            : ["Type = 'In-Process'"],
        logic: remark ? '{0} AND {1}' : '{0}'
    };
    JSUTIL.callAJAXPost('/data/OrderMaterial/list',
        JSON.stringify(payload),
        function (res) {
            const list = Array.isArray(res?.data) ? res.data : (Array.isArray(res) ? res : []);
            const ids = list.map(function (item) { return item.id || item.Id; }).filter(Boolean);
            ids.forEach(function (id) {
                JSUTIL.callAJAXPost(`/data/OrderMaterial/delete?id=${id}`,
                    '{}',
                    function () { },
                    function (err) {
                        console.error('Failed to delete target batch order material', err);
                    }
                );
            });
        },
        function (err) {
            console.error('Failed to fetch target batch order materials for deletion', err);
        }
    );
}

function renderProductTransferTable(list) {
    transactionsLoaded = list;
    const $tbody = $('#bpcrProductTransferTable');
    if (!$tbody.length) {
        return;
    }
    const batchNumber = (getBatchNumberForTransactions() || '').toString().trim().toLowerCase();
    const filtered = (Array.isArray(list) ? list : []).filter(function (txn) {
        if (!batchNumber) {
            return true;
        }
        const txnBatch = (txn.batchNumber || txn.BatchNumber || '').toString().trim().toLowerCase();
        const txnDoc = (txn.docNumber || txn.DocNumber || '').toString().trim().toLowerCase();
        return txnBatch === batchNumber || txnDoc === batchNumber;
    });
    if (!filtered.length) {
        $tbody.html('<tr><td colspan="8" class="text-center text-muted">No transactions.</td></tr>');
        return;
    }
    const sorted = filtered.slice().sort(function (a, b) {
        const aTime = new Date(a.transactionDate || a.TransactionDate || 0).getTime();
        const bTime = new Date(b.transactionDate || b.TransactionDate || 0).getTime();
        return aTime - bTime;
    });
    let runningBalance = 0;
    let totalReceipt = 0;
    let totalIssue = 0;
    const rows = sorted.map(function (txn) {
        const quantities = getTransferQuantities(txn);
        runningBalance += quantities.receiptValue - quantities.issueValue;
        totalReceipt += quantities.receiptValue;
        totalIssue += quantities.issueValue;
        const productInfo = getTransferProductInfo(txn);
        const productName = productInfo.name
            || finalProductItem?.product?.name
            || '';
        const productUom = productInfo.uom
            || finalProductItem?.product?.uom
            || '';
        const transferTo = txn.issuedToName || '';
        const txnDocCode = txn.batchNumber
            || txn.BatchNumber
            || txn.docNumber
            || txn.DocNumber
            || '';
        const txnId = txn.id || txn.Id || '';
        const changeReasonId = txn.changeReason || txn.ChangeReason || '';
        const safeTxnId = escapeHtml(txnId);
        const safeDocCode = escapeHtml(txnDocCode);
        const deleteBtn = `<button type="button" class="btn btn-xs btn-danger delete-transfer-btn"
                title="Delete" data-txn-id="${safeTxnId}" data-change-reason="${escapeHtml(changeReasonId)}"
                data-doc-code="${safeDocCode}">
                <i class="fa fa-trash"></i>
            </button>`;
        const actionButtons = [deleteBtn].filter(Boolean).join('');
        return `
            <tr>
                <td>${formatDateShortDisplay(txn.transactionDate || txn.TransactionDate)}</td>
                <td>${txnDocCode}</td>
                <td>${transferTo || '-'}</td>
                <td>${productName || '-'}</td>
                <td class="text-right">${formatQtyWithUom(quantities.receiptValue, productUom)}</td>
                <td class="text-right">${formatQtyWithUom(quantities.issueValue, productUom)}</td>
                <td class="text-right">${formatQtyWithUom(runningBalance, productUom)}</td>
                <td class="text-center">${actionButtons || '-'}</td>
            </tr>
        `;
    }).join('');
    const summaryUom = finalProductItem?.product?.uom || getTransferProductInfo(sorted[0]).uom || '';
    const summaryRow = `
        <tr class="active">
            <td colspan="4" class="text-right"><strong>Totals</strong></td>
            <td class="text-right"><strong>${formatQtyWithUom(totalReceipt, summaryUom)}</strong></td>
            <td class="text-right"><strong>${formatQtyWithUom(totalIssue, summaryUom)}</strong></td>
            <td class="text-right"><strong>${formatQtyWithUom(runningBalance, summaryUom)}</strong></td>
            <td></td>
        </tr>
    `;
    $tbody.html(rows + summaryRow);
}

function getTransferQuantities(transaction) {
    const qtyReceipt = Number(transaction.quantityReceipt || transaction.QuantityReceipt || 0);
    const qtyIssued = Number(transaction.quantityIssued || transaction.QuantityIssued || 0);
    const qtyChange = Number(transaction.quantityChange || transaction.QuantityChange || 0);
    let receiptValue = 0;
    let issueValue = 0;
    if (qtyReceipt || qtyIssued) {
        receiptValue = qtyReceipt || 0;
        issueValue = qtyIssued || 0;
    } else if (qtyChange > 0) {
        receiptValue = qtyChange;
    } else if (qtyChange < 0) {
        issueValue = Math.abs(qtyChange);
    }
    return {
        receiptValue: receiptValue,
        issueValue: issueValue,
        receipt: receiptValue ? formatQtyValue(receiptValue) : '',
        issue: issueValue ? formatQtyValue(issueValue) : ''
    };
}

function getTransferProductInfo(transaction) {
    const directProduct = transaction?.product || transaction?.Product || {};
    const nestedProduct = transaction?.orderMaterial?.product
        || transaction?.orderMaterial?.Product
        || transaction?.OrderMaterial?.Product
        || {};
    const product = Object.keys(directProduct).length ? directProduct : nestedProduct;
    return {
        id: product.id || product.Id || '',
        name: product.name || product.Name || '',
        number: product.number || product.Number || '',
        uom: product.uom || product.Uom || ''
    };
}

function formatQtyValue(value) {
    if (!Number.isFinite(value)) {
        return '';
    }
    const rounded = Math.round(value * 1000) / 1000;
    return Number.isInteger(rounded) ? String(rounded) : rounded.toFixed(2);
}

function formatQtyWithUom(value, uom) {
    const qty = formatQtyValue(value);
    if (!qty) {
        return '';
    }
    return uom ? `${qty} ${uom}` : qty;
}

function formatDateShortDisplay(dateString) {
    if (!dateString) {
        return '';
    }
    const date = new Date(dateString);
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
}

function toggleFinishProductionButton(status) {
    const showFinish = shouldShowFinishProductionButton(status);
    $('.finish-production-btn').toggle(showFinish);
}

function shouldShowFinishProductionButton(status) {
    const normalized = (status || '').toLowerCase();
    return normalized === 'rm-request'
        || normalized === 'rm-requested'
        || normalized === 'rm request'
        || normalized === 'rm requested'
        || normalized === 'rm issued'
        || normalized === 'rm-issued';
}

function handleFinishProduction() {
    if (!currentBPCRId) {
        return;
    }
    const batchNumber = getBatchNumberForTransactions();
    const quantityProduced = firstDefinedNumber(finalProductItem?.plannedQuantity, $('#finalProductQty').val(), 0);
    const orderMaterialId = finalProductItem?.id || '';
    const productId = finalProductItem?.product?.id || finalProductItem?.product?.Id || '';
    const docNumber = ($('#finalProductDocNumber').val() || getBPCRDocNumber() || '').toString().trim();

    if (!quantityProduced || quantityProduced <= 0) {
        JSUTIL.buildErrorModal('Enter a valid quantity produced before finishing production.');
        return;
    }
    if (!batchNumber) {
        JSUTIL.buildErrorModal('Batch number missing. Please refresh and try again.');
        return;
    }

    const targetStatus = BPCR_PACKING_LIST_STATUS;
    const $finishBtn = $('.finish-production-btn');
    $finishBtn.prop('disabled', true).hide();
    createProductionOutputTransaction({
        docNumber: docNumber,
        batchNumber: batchNumber,
        quantity: quantityProduced,
        orderMaterialId: orderMaterialId,
        productId: productId
    }, function () {
        updateBPCRStatusValue(targetStatus, function () {
            currentStatusValue = targetStatus;
            refreshCurrentBPCRRecord(currentBPCRId);
            refreshProductTransferList();
            renderFinalProductSection(finalProductItem);
            updateProgressSteps(targetStatus);
            $finishBtn.hide();
        }, function () {
            $finishBtn.show().prop('disabled', false);
        });
    }, function () {
        $finishBtn.show().prop('disabled', false);
    });
    location.reload();
}

function createProductionOutputTransaction(options, onSuccess, onError) {
    const payload = {
        docNumber: options.docNumber || '',
        batchNumber: options.batchNumber || '',
        // type: 'Production Output',
        transactionDate: new Date().toISOString(),
        quantityChange: Number(options.quantity) || 0,
        type: "Received"
    };
    if (options.orderMaterialId) {
        payload.orderMaterial = { id: options.orderMaterialId };
    }

    console.log("payload : ", payload)
    JSUTIL.callAJAXPost('/data/OrderMaterialTransaction/create',
        JSON.stringify(payload),
        function (res) {
            if (typeof onSuccess === 'function') {
                onSuccess(res);
            }
        },
        function (err) {
            console.error('Failed to create product transfer transaction', err);
            if (typeof onError === 'function') {
                onError(err);
            }
        }
    );
}

function openAssignModal() {
    if (!currentBPCRId) {
        // JSUTIL.buildErrorModal'Select a BPCR before assigning.');
        return;
    }
    fetchEmployees(function (list) {
        initAssignTypeahead(list || []);
        $('#ownerName').val('');
        $('#assignBPCRHiddenId').val('');
        $('#assignBPCRModal').modal('show');
    });
}

function fetchEmployees(callback) {
    const payload = {
        fields: 'Name',
        // logic: '{0}'
    };
    JSUTIL.callAJAXPost('/data/Employee/list', JSON.stringify(payload), function (res) {
        if (typeof callback === 'function') {
            callback(res || []);
        }
    });
}

function submitAssign() {
    const employeeId = $('.selectSearchInputAutoCompleteId').val();
    if (!employeeId) {
        // JSUTIL.buildErrorModal'Please select an employee.');
        return;
    }
    const payload = {
        orderOwner: { id: Number(employeeId) },
        orderStatus: BPCR_ASSIGNED_STATUS
    };
    JSUTIL.callAJAXPost(`/data/ProductionOrder/update/${currentBPCRId}`,
        JSON.stringify(payload),
        function () {
            // JSUTIL.buildErrorModal'BPCR assigned.');
            if (currentBPCRRecord) {
                currentBPCRRecord.orderStatus = BPCR_ASSIGNED_STATUS;
                currentBPCRRecord.orderOwner = { id: Number(employeeId) };
            }
            $('#assignBPCRModal').modal('hide');
            populateBPCRHeader(currentBPCRRecord, currentBPCRId);
            location.reload()
        }
    );
}

function initAssignTypeahead(list) {
    const mapped = (list || []).map(function (item) {
        const name = item.name || item.Name || '';
        return {
            id: item.id || item.Id,
            text: name,
            search: (name + ' ' + (item.id || item.Id || '')).toLowerCase()
        };
    });
    try {
        $('#ownerName').typeahead('destroy');
    } catch (e) {
        // ignore
    }
    $('#ownerName').typeahead({
        source: function (query, process) {
            query = (query || '').toLowerCase();
            const results = mapped
                .filter(function (i) { return i.search.indexOf(query) !== -1; })
                .map(function (i) { return i.text; });
            process(results);
        },
        afterSelect: function (selectedText) {
            const chosen = mapped.find(function (i) { return i.text === selectedText; });
            $('#assignBPCRHiddenId').val(chosen ? chosen.id : '');
        }
    });
}

function showActualTimeError(message) {
    const text = message || 'Enter valid start and stop times.';
    if (typeof JSUTIL !== 'undefined' && typeof JSUTIL.buildErrorModal === 'function') {
        JSUTIL.buildErrorModal(text);
    } else {
        JSUTIL.buildErrorModal(text);
    }
}

function showMaterialRequestError(message) {
    const text = message || 'Enter a requested quantity for at least one material.';
    if (typeof JSUTIL !== 'undefined' && typeof JSUTIL.buildErrorModal === 'function') {
        JSUTIL.buildErrorModal(text);
    } else {
        JSUTIL.buildErrorModal(text);
    }
}

function showFinalProductError(message) {
    const text = message || 'Enter valid final product details.';
    if (typeof JSUTIL !== 'undefined' && typeof JSUTIL.buildErrorModal === 'function') {
        JSUTIL.buildErrorModal(text);
    } else {
        JSUTIL.buildErrorModal(text);
    }
}

function getRowMaterialType($row) {
    const typeValue = $row ? $row.data('material-type') : '';
    if (typeValue) {
        return getNormalizedMaterialType(typeValue);
    }
    const tbodyId = ($row && $row.closest('tbody').attr('id') || '').toLowerCase();
    if (tbodyId.indexOf('raw') !== -1) {
        return 'raw';
    }
    if (tbodyId.indexOf('pack') !== -1) {
        return 'packing';
    }
    if (tbodyId.indexOf('general') !== -1) {
        return 'general';
    }
    if (tbodyId.indexOf('process') !== -1) {
        return 'inprocess';
    }
    return '';
}

function collectMaterialRequestEntries() {
    const $filter = $('#bpcrMaterialTypeFilter');
    const selectedValue = $filter.length ? $filter.val() : '';
    const $selectedOption = $filter.find('option').filter(function () { return $(this).val() === selectedValue; }).first();
    const isStepFilter = ($selectedOption.data('filter-type') || '').toString() === 'step';
    const selectedType = getNormalizedMaterialType(selectedValue);
    const includeAll = isStepFilter || !selectedType || selectedType === MATERIAL_TYPE_FILTER_ALL;
    const rows = $('#editBPCRPanel tbody tr');
    const materials = [];
    rows.each(function () {
        const $row = $(this);
        if ($row.data('in-process')) {
            return;
        }
        if (isStepFilter && !$row.is(':visible')) {
            return;
        }
        const rowType = getRowMaterialType($row);
        if (!includeAll) {
            if (!rowType || rowType !== selectedType) {
                return;
            }
        }
        const productId = $row.data('product-id');
        const orderMaterialId = $row.data('id');
        const productType = getProductTypeFromRow($row);
        const plannedQty = Number($row.data('planned-qty')) || 0;
        const uomText = ($row.data('uom') || '').toString().trim();
        if (!productId || !orderMaterialId) {
            return;
        }
        const inputVal = $row.find('.bpcr-request-qty-input').val();
        const parsedQty = Number(inputVal);
        const requestedQty = Number.isFinite(parsedQty) && parsedQty >= 0 ? parsedQty : 0;
        materials.push({
            orderMaterialId: orderMaterialId,
            productId: productId,
            productType: productType,
            requestedQty: requestedQty,
            plannedQty: plannedQty,
            uom: uomText
        });
    });
    return materials;
}

function getProductTypeFromRow($row) {
    if (!$row || !$row.length) {
        return '';
    }
    const rowType = ($row.data('product-type') || '').toString().trim();
    if (rowType) {
        return rowType;
    }
    const orderMaterialId = $row.data('id');
    if (!orderMaterialId) {
        return '';
    }
    const match = (bpcrMaterialItems || []).find(function (item) {
        return String(item.id) === String(orderMaterialId);
    });
    const product = match?.product || {};
    return (product.type || product.Type || '').toString().trim();
}

function deriveProductRequestType(materials) {
    const types = (materials || []).map(function (m) {
        return (m.productType || '').toString().trim();
    }).filter(Boolean);
    if (!types.length) {
        return '';
    }
    const unique = Array.from(new Set(types));
    return unique.length === 1 ? unique[0] : 'Mixed';
}

function collectUsedEntries() {
    const rows = $('#editBPCRPanel tbody tr');
    const materials = [];
    rows.each(function () {
        const $row = $(this);
        if ($row.data('in-process')) {
            return;
        }
        const orderMaterialId = $row.data('id');
        const productId = $row.data('product-id');
        const productName = ($row.find('td').eq(0).text() || '').trim();
        const rowLot = ($row.data('step-name') || $row.data('step-label') || '').toString();
        const issuedQty = getIssuedQuantityForProduct(productId, rowLot);
        if (!orderMaterialId) {
            return;
        }
        const inputVal = $row.find('.bpcr-used-qty-input').val();
        const parsedQty = Number(inputVal);
        const usedQty = Number.isFinite(parsedQty) ? parsedQty : 0;
        materials.push({
            orderMaterialId: orderMaterialId,
            quantityConsumed: usedQty,
            productId: productId,
            productName: productName,
            issuedQty: issuedQty
        });
    });
    return materials;
}

function validateUsedEntries(materials) {
    for (let i = 0; i < materials.length; i += 1) {
        const m = materials[i];
        const issued = Number.isFinite(m.issuedQty) ? m.issuedQty : 0;
        const used = Number.isFinite(m.quantityConsumed) ? m.quantityConsumed : 0;
        if (used > issued) {
            const name = (m.productName || 'material').trim();
            return { valid: false, message: `Used quantity for ${name} cannot exceed issued quantity (${issued}).` };
        }
        if (used < 0) {
            const name = (m.productName || 'material').trim();
            return { valid: false, message: `Used quantity for ${name} cannot be negative.` };
        }
    }
    return { valid: true };
}

function validateMaterialRequestEntries(materials) {
    if (!materials.length) {
        return { valid: false, message: 'Enter a requested quantity for at least one material.' };
    }

    // Removed: validation for requestedQty > 0

    return { valid: true };
}


function syncLocalRequestedQuantities(materials) {
    const qtyMap = {};
    materials.forEach(function (m) {
        qtyMap[String(m.orderMaterialId)] = m.requestedQty;
        const $row = $(`#editBPCRPanel tbody tr[data-id="${m.orderMaterialId}"]`);
        if ($row && $row.length) {
            $row.attr('data-request-qty', m.requestedQty);
        }
    });
    bpcrMaterialItems = (bpcrMaterialItems || []).map(function (item) {
        if (!item || !item.id) {
            return item;
        }
        const key = String(item.id);
        if (Object.prototype.hasOwnProperty.call(qtyMap, key)) {
            return { ...item, actualQuantity: qtyMap[key] };
        }
        return item;
    });
}

function persistRequestedMaterialQuantities(materials, onSuccess, onError) {
    const payload = materials.map(function (m) {
        return {
            id: m.orderMaterialId,
            actualQuantity: m.requestedQty
        };
    }).filter(function (item) {
        return item.id;
    });
    if (!payload.length) {
        if (typeof onSuccess === 'function') {
            onSuccess();
        }
        return;
    }
    JSUTIL.callAJAXPost('/data/OrderMaterial/upsert_multiple',
        JSON.stringify(payload),
        function (res) {
            syncLocalRequestedQuantities(materials);
            if (typeof onSuccess === 'function') {
                onSuccess(res);
            }
        },
        function (err) {
            console.error('Failed to save requested quantities', err);
            if (typeof onError === 'function') {
                onError(err);
            }
        }
    );
}

function persistUsedQuantities(materials, onSuccess, onError) {
    const payload = materials.map(function (m) {
        return {
            id: m.orderMaterialId,
            quantityConsumed: m.quantityConsumed
        };
    }).filter(function (item) { return item.id; });
    if (!payload.length) {
        if (typeof onSuccess === 'function') {
            onSuccess();
        }
        return;
    }
    JSUTIL.callAJAXPost('/data/OrderMaterial/upsert_multiple',
        JSON.stringify(payload),
        function (res) {
            materials.forEach(function (m) {
                const key = String(m.orderMaterialId);
                bpcrMaterialItems = (bpcrMaterialItems || []).map(function (item) {
                    if (String(item.id) === key) {
                        return { ...item, quantityConsumed: m.quantityConsumed };
                    }
                    return item;
                });
            });
            refreshReturnedQtyForAllRows();
            if (typeof onSuccess === 'function') {
                onSuccess(res);
            }
        },
        function (err) {
            console.error('Failed to save used quantities', err);
            if (typeof onError === 'function') {
                onError(err);
            }
        }
    );
}

function submitProductRequest(materials, lotLabel, onSuccess, onError) {
    const requestPayload = {
        // type: 'BPCR',
        processStatus: 'Requested',
        batchNumber: currentBPCRRecord?.name || '',
        // workOrder: { id: Number(currentBPCRId) },
        // remarks: `BPCR: ${currentBPCRRecord?.number || currentBPCRId}`
    };
    const requestType = deriveProductRequestType(materials);
    if (requestType) {
        requestPayload.type = requestType;
    }
    const remarksText = buildProductRequestRemark(currentBPCRRecord, currentBPCRId, lotLabel);
    if (remarksText) {
        requestPayload.remarks = remarksText;
    }
    if (typeof employeeId !== 'undefined' && employeeId) {
        requestPayload.requestedBy = { id: Number(employeeId) };
        if (typeof employeeName !== 'undefined') {
            requestPayload.requestedByName = employeeName;
        }
    }
    else {
        JSUTIL.buildErrorModal("Your user is not assigned to an employee, contact administrator for help.")
        return
    }

    JSUTIL.callAJAXPost('/data/ProductRequest/create_with_sequence?field=Number',
        JSON.stringify(requestPayload),
        function (reqRes) {
            const reqId = reqRes?.id;
            if (!reqId) {
                if (typeof onError === 'function') {
                    onError();
                }
                return;
            }
            const itemPayload = materials.map(function (m) {
                return {
                    productRequest: { id: reqId },
                    product: { id: Number(m.productId) },
                    quantityRequested: m.requestedQty,
                    quantityIssued: 0,
                    quantityRemaining: m.requestedQty,
                    quantityRemainingUom: m.uom || ''
                };
            });
            JSUTIL.callAJAXPost('/data/ProductRequestItem/upsert_multiple',
                JSON.stringify(itemPayload),
                function (res) {
                    if (typeof onSuccess === 'function') {
                        onSuccess(res);
                    }
                },
                function (err) {
                    console.error('Failed to create ProductRequestItem', err);
                    if (typeof onError === 'function') {
                        onError(err);
                    }
                }
            );
        },
        function (err) {
            console.error('Failed to create ProductRequest', err);
            if (typeof onError === 'function') {
                onError(err);
            }
        }
    );
}

function prepareMaterialRequest() {
    if (!currentBPCRId) {
        // JSUTIL.buildErrorModal'Select a BPCR first.');
        return;
    }
    const lotInfo = resolveLotForRequest();
    if (lotInfo.requiresSelection) {
        showMaterialRequestError('Select a lot before sending the material request.');
        return;
    }
    const lotLabel = lotInfo.label || lotInfo.normalized || '';
    const lotAlreadyRequested = lotInfo.normalized ? isLotAlreadyRequested(lotInfo.normalized) : false;
    if (lotAlreadyRequested) {
        showMaterialRequestError('A material request has already been sent for this product lot.');
        return;
    }
    const materials = collectMaterialRequestEntries();
    const validation = validateMaterialRequestEntries(materials);
    if (!validation.valid) {
        showMaterialRequestError(validation.message);
        return;
    }
    $('.request-materials').prop('disabled', true);

    const resetRequestButton = function () {
        $('.request-materials').prop('disabled', false);
        location.reload();
    };

    persistRequestedMaterialQuantities(materials, function () {
        submitProductRequest(materials, lotLabel, function () {
            hasExistingMaterialRequest = true;
            if (lotLabel) {
                markRequestedLot(lotLabel);
            }
            refreshRequestQtyInputsReadonly();
            toggleAssignButton(currentStatusValue);
            updateBPCRStatusValue('RM-Request', function () {
                resetRequestButton();
            }, function () {
                resetRequestButton();
            });
        }, function () {
            resetRequestButton();
        });
    }, function () {
        resetRequestButton();
    });
}

function saveUsedQuantities() {
    const materials = collectUsedEntries().filter(function (m) { return Number.isFinite(m.quantityConsumed); });
    if (!materials.length) {
        showMaterialRequestError('Enter used quantity for at least one material.');
        return;
    }
    const validation = validateUsedEntries(materials);
    if (!validation.valid) {
        showMaterialRequestError(validation.message);
        return;
    }
    $('.save-used-qty').prop('disabled', true);
    persistUsedQuantities(materials, function () {
        refreshReturnedQtyForAllRows();
        $('.save-used-qty').prop('disabled', false);
    }, function () {
        $('.save-used-qty').prop('disabled', false);
    });
}

function buildProductRequestRemark(record, fallbackId, lotLabel) {
    const orderNumber = (record?.number || '').toString().trim();
    const idPart = fallbackId ? ` (ID:${fallbackId})` : '';
    const baseText = `${orderNumber || fallbackId || ''}${idPart}`.trim();
    let remark = baseText;
    if (lotLabel) {
        remark = remark ? `${remark} | Lot: ${lotLabel}` : `Lot: ${lotLabel}`;
    }
    return remark.trim();
}

function updateBPCRStatusValue(statusValue, onSuccess, onError) {
    if (!currentBPCRId || !statusValue) {
        if (typeof onSuccess === 'function') {
            onSuccess();
        }
        return;
    }
    const payload = { orderStatus: statusValue };
    JSUTIL.callAJAXPost(`/data/ProductionOrder/update/${currentBPCRId}`,
        JSON.stringify(payload),
        function (res) {
            if (currentBPCRRecord) {
                currentBPCRRecord.orderStatus = statusValue;
            }
            populateBPCRHeader(currentBPCRRecord, currentBPCRId);
            if (typeof onSuccess === 'function') {
                onSuccess(res);
            }
        },
        function (err) {
            console.error('Failed to update BPCR status', err);
            if (typeof onError === 'function') {
                onError(err);
            }
        }
    );
}

function refreshCurrentBPCRRecord(orderIdParam) {
    const orderId = orderIdParam || currentBPCRId;
    if (!orderId) {
        return;
    }
    const payload = {
        fields: 'Id;ActualStartTime;ActualEndTime;OrderStatus',
        conditions: [`Id = ${orderId}`],
        logic: '{0}'
    };
    JSUTIL.callAJAXPost('/data/ProductionOrder/list',
        JSON.stringify(payload),
        function (res) {
            const list = Array.isArray(res?.data)
                ? res.data
                : (Array.isArray(res) ? res : []);
            if (!list || !list.length) {
                return;
            }
            const record = list[0] || {};
            if (!currentBPCRRecord) {
                currentBPCRRecord = {};
            }
            currentBPCRRecord.actualStartTime = record.actualStartTime
                || record.ActualStartTime
                || '';
            currentBPCRRecord.actualEndTime = record.actualEndTime
                || record.ActualEndTime
                || '';
            if (record.orderStatus || record.OrderStatus) {
                currentBPCRRecord.orderStatus = record.orderStatus || record.OrderStatus;
            }
            populateBPCRHeader(currentBPCRRecord, orderId);
        },
        function (err) {
            console.error('Failed to refresh BPCR record', err);
        }
    );
}
function loadBPCRDetails(orderId) {
    intermediateAddedTxnIds.clear();

    let payload = {
        fields: 'Type;PlannedQuantity;ActualQuantity;QuantityConsumed;NumberOfContainers;Product.Id;Remarks;DocNumber;ManufacturingDate;ExpiryDate;RetestDate;ProductLot;ExpiryType;Grade;Variation',
        conditions: [`Remarks = '${orderId}'`],
        logic: "{0}"
    };

    JSUTIL.callAJAXPost('/data/OrderMaterial/list', JSON.stringify(payload), function (items) {

        console.log("BOM Materials:", items);
        const listResponse = Array.isArray(items?.data) ? items.data : (items || []);
        const normalizedList = listResponse.map(function (item) {
            return normalizeMaterialItem(item);
        });
        const finalItems = normalizedList.filter(function (i) {
            return isFinalProductEntry(i);
        });
        finalProductItem = selectFinalProductItem(finalItems);
        bpcrMaterialItems = normalizedList.filter(function (i) {
            if (isBsrOrderMaterial(i)) {
                return false;
            }
            return !isFinalProductEntry(i);
        });
        bpcrIssuedQuantityMap = {};
        bpcrRequestedQuantityMap = {};
        issueRequestItems = [];
        renderBPCRMaterialTables(bpcrMaterialItems);
        renderFinalProductSection(finalProductItem);
        packingListLoadedFromServer = false;
        packingListFetchInProgress = false;
        fetchIssuedMaterialQuantities(orderId, function (issuedMap, details, requestedMap) {
            bpcrIssuedQuantityMap = issuedMap || {};
            bpcrRequestedQuantityMap = requestedMap || {};
            issueRequestItems = Array.isArray(details) ? details : [];
            renderBPCRMaterialTables(bpcrMaterialItems);
            renderIssueRequestTable(issueRequestItems);
        });
        refreshProductTransferList();
    }, function (err) {
        console.error('Failed to load BPCR materials', err);
        materialRequestCheckPending = false;
        toggleAssignButton(currentStatusValue);
    });
}
function renderBPCRMaterialTables(items) {

    const showIssued = shouldShowIssuedColumn();
    const showRequested = shouldShowRequestedColumn();
    const showUsed = shouldShowUsedColumn();
    const showReturned = showUsed;
    const requestLabel = getRequestColumnLabel();
    const groups = groupBPCRMaterials(items || []);

    buildBPCRMaterialRows('#bpcrRawMaterialTable', groups.raw, showIssued, showRequested, showUsed, showReturned, requestLabel, 'raw');
    buildBPCRMaterialRows('#bpcrPackingMaterialTable', groups.packing, showIssued, showRequested, showUsed, showReturned, requestLabel, 'packing');
    buildBPCRMaterialRows('#bpcrGeneralMaterialTable', groups.general, showIssued, showRequested, showUsed, showReturned, requestLabel, 'general');
    buildInProcessMaterialRows('#bpcrInProcessMaterialTable', groups.inProcess);
    toggleIssuedQtyHeaders(showIssued);
    toggleRequestedQtyHeaders(showRequested, requestLabel);
    toggleUsedQtyHeaders(showUsed);
    toggleReturnedQtyHeaders(showReturned);
    refreshMaterialTypeFilterOptions(items);
    applyMaterialTypeFilter();
}

function renderIssueRequestTable(list) {
    const $container = $('#issueRequestTableContainer');
    if (!$container.length) {
        return;
    }
    if (!Array.isArray(list) || !list.length) {
        $container.html('<div class="text-center text-muted">No requests yet.</div>');
        return;
    }
    const grouped = {};
    list.forEach(function (item, idx) {
        const key = item.requestNumber || `PI-${item.requestId || idx + 1}`;
        if (!grouped[key]) {
            grouped[key] = [];
        }
        grouped[key].push(item);
    });
    const tables = Object.keys(grouped).sort().map(function (reqKey) {
        const firstItem = grouped[reqKey][0] || {};
        const reqStatus = firstItem?.status || '';
        const normalizedStatus = (reqStatus || '').toString().toLowerCase();
        const requestId = firstItem?.requestId || '';
        const hasIssued = grouped[reqKey].some(function (item) {
            return Number(item.quantityIssued || 0) > 0;
        });
        const canCancel = requestId
            && !hasIssued
            && normalizedStatus !== 'issued'
            && !isCancelledRequestStatus(reqStatus);
        const cancelButton = canCancel
  ? `
    <button
      class="btn btn-danger btn-sm cancel-issue-request-btn"
      style="margin-left:auto; display:block;"
      data-request-id="${escapeHtml(requestId)}"
      data-request-number="${escapeHtml(reqKey)}">
      Cancel Request
    </button>
  `
  : '';


        const itemRows = grouped[reqKey].map(function (item) {
            const productText = [item.productName, item.productCode].filter(Boolean).join(' | ') || '-';
            const requested = Number(item.quantityRequested || 0);
            const issued = Number(item.quantityIssued || 0);
            const remarksText = '';
            return `
                <tr>
                    <td>${formatPrintValue(productText)}</td>
                    <td>${requested}</td>
                    <td>${issued}</td>
                </tr>
            `;
        }).join('');
        return `
            <div class="issue-request-block" style="margin-bottom:16px;">
                <div class="section-header clearfix">
                    
                    <h6 class="details-subtitle">${formatPrintValue(reqKey)}${reqStatus ? ' - ' + formatPrintValue(reqStatus) : ''}</h6>
                    ${cancelButton}
                </div>
                <div class="table-responsive">
                    <table class="table table-bordered table-hover">
                        <thead>
                            <tr>
                                <th>Material</th>
                                <th style="width:130px;">Requested Qty</th>
                                <th style="width:130px;">Issued Qty</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${itemRows}
                        </tbody>
                    </table>
                </div>
            </div>
        `;
    }).join('');
    $container.html(tables);
}

function handleIssueRequestCancel($button) {
    if (!$button || !$button.length) {
        return;
    }
    const requestId = $button.data('requestId');
    const requestNumber = $button.data('requestNumber') || '';
    if (!requestId) {
        return;
    }
    const promptText = requestNumber
        ? `Cancel request ${requestNumber}?`
        : 'Cancel this request?';
    if (!confirm(promptText)) {
        return;
    }
    const originalText = $button.text();
    $button.prop('disabled', true).text('Cancelling...');
    JSUTIL.callAJAXPost(`/data/ProductRequest/update/${requestId}`,
        JSON.stringify({ processStatus: 'Cancelled' }),
        function () {
            issueRequestItems = (issueRequestItems || []).map(function (item) {
                if (String(item.requestId) === String(requestId)) {
                    return { ...item, status: 'Cancelled' };
                }
                return item;
            });
            refreshRequestedLotsFromIssueRequests(issueRequestItems);
            renderIssueRequestTable(issueRequestItems);
            JSUTIL.buildErrorModal('Request cancelled.');
        },
        function (err) {
            console.error('Failed to cancel product request', err);
            JSUTIL.buildErrorModal('Could not cancel this request. Please try again.');
            $button.prop('disabled', false).text(originalText);
        }
    );
}

function groupBPCRMaterials(items) {
    const groups = {
        raw: [],
        packing: [],
        general: [],
        inProcess: []
    };
    (items || []).forEach(function (item) {
        const typeValue = getMaterialTypeValue(item);
        const normalized = typeValue.replace(/\s+/g, '');
        if (normalized.indexOf('in-process') !== -1 || normalized === 'inprocess') {
            groups.inProcess.push(item);
        } else if (typeValue.indexOf('raw') !== -1) {
            groups.raw.push(item);
        } else if (typeValue.indexOf('pack') !== -1) {
            groups.packing.push(item);
        } else {
            groups.general.push(item);
        }
    });
    return groups;
}

function renderFinalProductSection(item) {
    const $container = $('#bpcrFinalProductContent');
    if (!$container.length) {
        return;
    }
    const isFinalStage = isFinalStageForCurrentBatch();
    const showFinishButton = shouldShowFinishProductionButton(currentStatusValue);
    const finishButtonHtml = showFinishButton
        ? `<button type="button" class="btn btn-warning btn-sm finish-production-btn">
                                            <i class="fa fa-flag-checkered"></i> Finish
                                        </button>`
        : '';
    const productName = item?.product?.name
        || currentBPCRRecord?.productRecipe?.productStage?.product?.name
        || '';
    const productCode = item?.product?.number
        || currentBPCRRecord?.productRecipe?.productStage?.product?.number
        || '';
    const uom = item?.product?.uom
        || currentBPCRRecord?.productRecipe?.productStage?.product?.uom
        || '';
    const quantity = Number(item?.plannedQuantity || 0) || '';
    const id = item?.id || '';
    const productId = item?.product?.id || item?.product?.Id
        || currentBPCRRecord?.productRecipe?.productStage?.product?.id
        || currentBPCRRecord?.productRecipe?.product?.id
        || '';
    const productInfo = [productName, productCode].filter(Boolean).join(' | ');
    const normalizedStatus = (currentStatusValue || '').toLowerCase();
    const isLocked = isFinalProductLockedStatus(currentStatusValue);
    const lockMessage = normalizedStatus.indexOf('under test') !== -1
        ? 'BPCR under test. Final product details are locked.'
        : (isCompletedStatus(currentStatusValue)
            ? 'BPCR completed. Final product details are locked.'
            : 'Production ended. Final product details are locked.');
    const readOnlyAttr = (isLocked || !finalProductEditEnabled) ? 'readonly disabled' : '';
    const docNumberValue = item?.docNumber
        || getBPCRDocNumber()
        || '';
    const netQtyPerContainer = Number.isFinite(Number(item?.actualQuantity))
        ? Number(item.actualQuantity)
        : '';
    const variationPerContainer = Number.isFinite(Number(item?.variation))
        ? Number(item.variation)
        : '';
    const gradeValue = item?.grade || '';
    const mfgDate = formatDateInputValue(item?.manufacturingDate || new Date());
    const expiryDate = formatDateInputValue(item?.expiryDate);
    const expiryType = normalizeExpiryType(item?.expiryType || 'Expiry');
    const expiryDateLabel = expiryType === 'Retest' ? 'Retest Date' : 'Expiry Date';
    const expiryTypeOptions = ['Expiry', 'Retest'];
    const expiryTypeOptionsHtml = expiryTypeOptions.map(function (type) {
        const selected = type === expiryType ? 'selected' : '';
        return `<option value="${type}" ${selected}>${type}</option>`;
    }).join('');
    const expiryTypeDisabled = (isLocked || !finalProductEditEnabled) ? 'disabled' : '';
    const buttonBar = `
        <div class="clearfix" style="margin-bottom:10px;">
            <div class="pull-right final-product-actions">
            ${finishButtonHtml}
                <button type="button" class="btn btn-success btn-sm final-product-edit-btn">
                    <i class="fa fa-pencil"></i> Edit
                </button>
                <button type="button" class="btn btn-success btn-sm final-product-save-btn" style="display:none;">
                    <i class="fa fa-check"></i> Save
                </button>
            </div>
        </div>`;

    const emptyText = (!productInfo && !productId)
        ? '<div class="text-muted final-product-empty-text">Start production to record the final product output.</div>'
        : '';

    let containerCountField = '';
    if (isFinalStage) {
        const numericContainerCount = Number(item?.numberOfContainers);
        const containerCountValue = Number.isFinite(numericContainerCount) && numericContainerCount > 0
            ? numericContainerCount
            : '';
        containerCountField = `
        <div class="form-group col-sm-6">
            <label>Number of Containers / Bags <span class="text-danger">*</span></label>
            <input type="number" class="form-control" id="finalProductNumberOfContainers" value="${containerCountValue}" min="1" step="1" ${readOnlyAttr} />
        </div>`;
    }

    const gradeOptions = ['', 'IP', 'USP'].map(function (opt) {
        const selected = opt === gradeValue ? 'selected' : '';
        const label = opt || 'Select';
        return `<option value="${opt}" ${selected}>${label}</option>`;
    }).join('');

    $container.html(`
        ${buttonBar}
        ${emptyText}
        <div class="form-group col-sm-3">
            <label>Quantity Produced <span class="text-danger">*</span></label>
            <input type="number" class="form-control" id="finalProductQty" value="${quantity}" min="0" step="any" ${readOnlyAttr} />
        </div>
        <div class="form-group col-sm-3">
            <label>UOM</label>
            <div class="bpcr-meta-value">${uom || '-'}</div>
        </div>
        ${containerCountField}
        <div class="form-group col-sm-6">
            <label>Net Quantity per Container</label>
            <input type="number" class="form-control" id="finalProductNetQtyPerContainer" value="${netQtyPerContainer}" min="0" step="any" ${readOnlyAttr} />
        </div>
        <div class="form-group col-sm-6">
            <label>Variation per Container</label>
            <input type="number" class="form-control" id="finalProductVariationPerContainer" value="${variationPerContainer}" step="any" ${readOnlyAttr} />
        </div>
        <div class="form-group col-sm-6">
            <label>Grade</label>
            <select class="form-control" id="finalProductGrade" ${readOnlyAttr}>
                ${gradeOptions}
            </select>
        </div>
        <div class="form-group col-sm-6">
            <label>Manufacturing Date <span class="text-danger">*</span></label>
            <input type="date" class="form-control datepicker" id="finalProductManufacturingDate" value="${mfgDate}" ${readOnlyAttr} />
        </div>
        <div class="form-group col-sm-6">
            <label>Expiry Type</label>
            <select class="form-control" id="finalProductExpiryType" ${expiryTypeDisabled}>
                ${expiryTypeOptionsHtml}
            </select>
        </div>
        <div class="form-group col-sm-6">
            <label id="finalProductExpiryDateLabel">${expiryDateLabel}</label>
            <input type="date" class="form-control datepicker" id="finalProductExpiryDate" value="${expiryDate}" ${readOnlyAttr} />
        </div>
        <input type="hidden" id="finalProductDocNumber" value="${docNumberValue}" />
        <input type="hidden" id="finalProductId" value="${id}" />
        <input type="hidden" id="finalProductProductId" value="${productId}" />
    `);
    updateFinalProductButtons(isLocked, !!productId);
    toggleFinalProductInputs(finalProductEditEnabled);
    updateFinalProductDateLabel();
    renderPackingListSection();
}

function getPackingListStorageKey() {
    return currentBPCRId ? `bpcrPackingList:${currentBPCRId}` : '';
}

function getPackingListContainerCount() {
    const savedCount = Number(finalProductItem?.numberOfContainers);
    if (Number.isFinite(savedCount) && savedCount > 0) {
        return savedCount;
    }
    const inputCount = Number($('#finalProductNumberOfContainers').val());
    if (Number.isFinite(inputCount) && inputCount > 0) {
        return inputCount;
    }
    return 0;
}

function normalizePackingListEntries(entries, containerCount) {
    const count = Number(containerCount) || 0;
    const list = Array.isArray(entries) ? entries : [];
    const normalized = [];
    for (let i = 0; i < count; i += 1) {
        const item = list[i] || {};
        const containerRaw = (item?.containerNumber ?? item?.containerNo ?? item?.number ?? item?.Number ?? '').toString().trim();
        const containerNumber = containerRaw || String(i + 1);
        normalized.push({
            id: item?.id || item?.Id || '',
            containerNumber: containerNumber,
            sealNumber: (item?.sealNumber ?? item?.SealNumber ?? '').toString().trim(),
            netWeight: Number.isFinite(Number(item?.netWeight ?? item?.NetWeight ?? item?.netWt)) ? Number(item.netWeight ?? item.NetWeight ?? item.netWt) : '',
            tareWeight: Number.isFinite(Number(item?.tareWeight ?? item?.TareWeight ?? item?.tareWt)) ? Number(item.tareWeight ?? item.TareWeight ?? item.tareWt) : '',
            grossWeight: Number.isFinite(Number(item?.grossWeight ?? item?.GrossWeight ?? item?.grossWt)) ? Number(item.grossWeight ?? item.GrossWeight ?? item.grossWt) : ''
        });
    }
    return normalized;
}

function loadPackingListEntries(containerCount) {
    const key = getPackingListStorageKey();
    let stored = [];
    if (key && typeof localStorage !== 'undefined') {
        try {
            const raw = localStorage.getItem(key);
            stored = raw ? JSON.parse(raw) : [];
        } catch (err) {
            console.warn('Could not parse stored packing list', err);
            stored = [];
        }
    }
    const normalized = normalizePackingListEntries(stored, containerCount);
    packingListEntries = normalized;
    return normalized;
}

function persistPackingListEntries() {
    const key = getPackingListStorageKey();
    if (!key || typeof localStorage === 'undefined') {
        return;
    }
    try {
        localStorage.setItem(key, JSON.stringify(packingListEntries || []));
    } catch (err) {
        console.warn('Could not persist packing list', err);
    }
}

function setPackingListMessage(message, isError) {
    const $msg = $('#bpcrPackingListMessage');
    if (!$msg.length) {
        return;
    }
    if (!message) {
        $msg.text('').hide();
        return;
    }
    $msg.removeClass('text-danger text-success')
        .addClass(isError ? 'text-danger' : 'text-success')
        .text(message)
        .show();
}

function setPackingListEditMode(enabled) {
    packingListEditEnabled = !!enabled && canEditPackingList(currentStatusValue);
    renderPackingListSection();
}

function togglePackingListTab(show) {
    const $tab = $('#bpcrPackingListTab');
    const $content = $('#bpcr-packing-list');
    const $materialTab = $('#bpcrMaterialTab');
    const $materialContent = $('#bpcr-material');
    if (!$tab.length || !$content.length) {
        return;
    }
    if (show) {
        $tab.show();
        return;
    }
    const isActive = $tab.hasClass('active') || $content.hasClass('active');
    $tab.removeClass('active').hide();
    $content.removeClass('active in');
    if (isActive && $materialTab.length && $materialContent.length) {
        $materialContent.addClass('active in');
        $materialTab.addClass('active');
    }
}

function getPackingListMeta() {
    const product = finalProductItem?.product
        || currentBPCRRecord?.productRecipe?.productStage?.product
        || currentBPCRRecord?.productRecipe?.product
        || {};
    const productInside = [product?.number, product?.name].filter(Boolean).join(' - ');
    const uom = product?.uom || '';
    const manufacturingDate = formatDateInputValue($('#finalProductManufacturingDate').val() || finalProductItem?.manufacturingDate);
    const expiryTypeRaw = $('#finalProductExpiryType').val() || finalProductItem?.expiryType || 'Expiry';
    const expiryType = normalizeExpiryType(expiryTypeRaw);
    const expiryDate = formatDateInputValue($('#finalProductExpiryDate').val() || finalProductItem?.expiryDate);
    const expiryLabel = expiryType === 'Retest' ? 'Retest Date' : 'Expiry Date';
    return {
        productInside: productInside,
        uom: uom,
        manufacturingDate: manufacturingDate,
        expiryDate: expiryDate,
        expiryLabel: expiryLabel,
        expiryType: expiryType
    };
}

function buildPackingListRows(entries, readOnly, containerCount) {
    if (!Array.isArray(entries) || !entries.length) {
        return '<tr><td colspan="9" class="text-center text-muted">No packing rows available.</td></tr>';
    }
    const attr = readOnly ? 'readonly disabled' : '';
    const formatValue = function (value) {
        if (value === 0 || value === '0') {
            return 0;
        }
        const num = Number(value);
        return Number.isNaN(num) ? '' : num;
    };
    const meta = getPackingListMeta();
    const total = Number(containerCount) || entries.length;
    return entries.map(function (entry, idx) {
        const rowIndex = idx + 1;
        const containerValue = total ? `${rowIndex} of ${total}` : String(rowIndex);
        const sealNumber = (entry?.sealNumber ?? '').toString();
        return `
            <tr data-row-index="${rowIndex}" data-id="${entry.id || ''}" data-container="${escapeHtml(containerValue)}">
                <td>${escapeHtml(containerValue)}</td>
                <td><input type="text" class="form-control input-sm packing-seal-input" value="${escapeHtml(sealNumber)}" ${attr}></td>
                <td>${escapeHtml(meta.productInside || '-')}</td>
                <td><input type="number" class="form-control input-sm packing-net-input" step="any" min="0" value="${formatValue(entry.netWeight)}" ${attr}></td>
                <td><input type="number" class="form-control input-sm packing-tare-input" step="any" min="0" value="${formatValue(entry.tareWeight)}" ${attr}></td>
                <td><input type="number" class="form-control input-sm packing-gross-input" step="any" min="0" value="${formatValue(entry.grossWeight)}" ${attr}></td>
                <td>${escapeHtml(meta.uom || '-')}</td>
                <td>${escapeHtml(meta.manufacturingDate || '-')}</td>
                <td>${escapeHtml(meta.expiryDate || '-')}</td>
            </tr>
        `;
    }).join('');
}

function renderPackingListSection() {
    const $tab = $('#bpcrPackingListTab');
    const $content = $('#bpcr-packing-list');
    const $body = $('#bpcrPackingListTableBody');
    const $editBtn = $('.packing-list-edit-btn');
    const $saveBtn = $('.packing-list-save-btn');
    const $expiryHeader = $('#bpcrPackingExpiryHeader');
    if (!$tab.length || !$content.length || !$body.length) {
        return;
    }
    const isFinalStage = isFinalStageForCurrentBatch();
    const statusAllowsPackingList = isPackingListVisibleStatus(currentStatusValue);
    const shouldShowTab = isFinalStage && statusAllowsPackingList;
    togglePackingListTab(shouldShowTab);
    if (!isFinalStage) {
        $body.html('<tr><td colspan="9" class="text-center text-muted">Packing list is available only for final stage batches.</td></tr>');
        $editBtn.prop('disabled', true).hide();
        $saveBtn.hide();
        setPackingListMessage('', false);
        return;
    }
    if (!statusAllowsPackingList) {
        $body.html('<tr><td colspan="9" class="text-center text-muted">Packing list is available after the status reaches Packing List Creation.</td></tr>');
        $editBtn.prop('disabled', true).hide();
        $saveBtn.hide();
        setPackingListMessage('', false);
        return;
    }
    const canEditStatus = canEditPackingList(currentStatusValue);
    if (!canEditStatus && packingListEditEnabled) {
        packingListEditEnabled = false;
    }
    const meta = getPackingListMeta();
    if ($expiryHeader && $expiryHeader.length) {
        $expiryHeader.text(meta.expiryLabel || 'Expiry / Retest Date');
    }
    const containerCount = getPackingListContainerCount();
    if (!Number.isFinite(containerCount) || containerCount <= 0) {
        $body.html('<tr><td colspan="9" class="text-center text-muted">Add the number of containers/bags in Output Product to generate packing rows.</td></tr>');
        $editBtn.prop('disabled', true).show();
        $saveBtn.hide();
        setPackingListMessage('Enter number of containers to create packing list rows.', false);
        return;
    }

    const entries = (packingListEntries && packingListEntries.length)
        ? normalizePackingListEntries(packingListEntries, containerCount)
        : loadPackingListEntries(containerCount);
    if (!packingListLoadedFromServer && !packingListFetchInProgress) {
        packingListFetchInProgress = true;
        fetchPackingListEntries(containerCount, function () {
            packingListFetchInProgress = false;
            packingListLoadedFromServer = true;
            renderPackingListSection();
        });
        return;
    }
    packingListEntries = normalizePackingListEntries(entries, containerCount);
    persistPackingListEntries();
    const allowEdit = canEditStatus && Number.isFinite(containerCount) && containerCount > 0;
    $body.html(buildPackingListRows(packingListEntries, !allowEdit || !packingListEditEnabled, containerCount));
    $editBtn.prop('disabled', !allowEdit).toggle(allowEdit && !packingListEditEnabled);
    $saveBtn.toggle(allowEdit && packingListEditEnabled);
    if (!allowEdit) {
        setPackingListMessage('Packing list is view only after Packing List Creation stage.', false);
        return;
    }
    setPackingListMessage('', false);
}

function syncPackingListWithContainerCount() {
    if (!isFinalStageForCurrentBatch()) {
        return;
    }
    const containerCount = getPackingListContainerCount();
    if (!Number.isFinite(containerCount) || containerCount <= 0) {
        return;
    }
    packingListEntries = normalizePackingListEntries(packingListEntries, containerCount);
    persistPackingListEntries();
    renderPackingListSection();
}

function fetchPackingListEntries(containerCount, onComplete) {
    if (!currentBPCRId) {
        if (typeof onComplete === 'function') {
            onComplete();
        }
        return;
    }
    const count = Number(containerCount) || 0;
    const payload = {
        fields: 'Id;Number;PacksPerLot;PackDescription;NetWeight;TareWeight;GrossWeight;Quantity;Remarks;Product.Id;SealNumber',
        conditions: [`Remarks = '${currentBPCRId}'`],
        logic: '{0}'
    };
    JSUTIL.callAJAXPost('/data/ProductContainer/list',
        JSON.stringify(payload),
        function (res) {
            const list = Array.isArray(res?.data) ? res.data : (Array.isArray(res) ? res : []);
            if (Array.isArray(list)) {
                const containers = normalizePackingListEntries(list, count || list.length || 0);
                packingListEntries = containers;
                persistPackingListEntries();
            }
            if (typeof onComplete === 'function') {
                onComplete();
            }
        },
        function (err) {
            console.error('Failed to load packing list', err);
            if (typeof onComplete === 'function') {
                onComplete();
            }
        }
    );
}

function buildPackingListPayload(entries) {
    const productId = finalProductItem?.product?.id
        || finalProductItem?.product?.Id
        || currentBPCRRecord?.productRecipe?.productStage?.product?.id
        || currentBPCRRecord?.productRecipe?.product?.id
        || '';
    const parseNumber = function (value) {
        const num = Number(value);
        return Number.isFinite(num) ? num : null;
    };
    return (entries || []).map(function (entry) {
        const sealNumber = (entry?.sealNumber ?? '').toString().trim();
        const payload = {
            id: entry.id || undefined,
            number: entry.containerNumber || '',
            netWeight: parseNumber(entry.netWeight),
            tareWeight: parseNumber(entry.tareWeight),
            grossWeight: parseNumber(entry.grossWeight),
            sealNumber: sealNumber || null,
            processStatus: 'New',
            remarks: currentBPCRId || '',
            status: true
        };
        if (entry.id) {
            payload.id = entry.id;
        }
        if (productId) {
            payload.product = { id: Number(productId) };
        }
        return payload;
    });
}

function savePackingListDetails() {
    if (!isFinalStageForCurrentBatch()) {
        setPackingListMessage('Packing list is only available for final stage batches.', true);
        return;
    }
    if (!canEditPackingList(currentStatusValue)) {
        setPackingListMessage('Packing list can only be edited in the Packing List Creation status.', true);
        return;
    }
    const containerCount = getPackingListContainerCount();
    if (!Number.isFinite(containerCount) || containerCount <= 0) {
        setPackingListMessage('Add number of containers/bags before saving packing list.', true);
        return;
    }
    const rows = [];
    let hasError = false;
    $('#bpcrPackingListTableBody tr').each(function () {
        const $row = $(this);
        const entryId = $row.data('id');
        const rowIndex = Number($row.data('rowIndex')) || rows.length + 1;
        const containerNumberRaw = ($row.data('container') || '').toString().trim();
        const containerNumber = containerNumberRaw
            || (Number.isFinite(containerCount) && containerCount > 0 ? `${rowIndex} of ${containerCount}` : String(rowIndex));
        if (!containerNumber) {
            hasError = true;
            return false;
        }
        const sealNumber = ($row.find('.packing-seal-input').val() || '').toString().trim();
        const netWeightRaw = $row.find('.packing-net-input').val();
        const tareWeightRaw = $row.find('.packing-tare-input').val();
        const grossWeightRaw = $row.find('.packing-gross-input').val();
        const netWeight = netWeightRaw === '' ? '' : Number(netWeightRaw);
        const tareWeight = tareWeightRaw === '' ? '' : Number(tareWeightRaw);
        const grossWeight = grossWeightRaw === '' ? '' : Number(grossWeightRaw);
        if ((netWeightRaw !== '' && !Number.isFinite(netWeight))
            || (tareWeightRaw !== '' && !Number.isFinite(tareWeight))
            || (grossWeightRaw !== '' && !Number.isFinite(grossWeight))) {
            hasError = true;
            return false;
        }
        rows.push({
            id: entryId || '',
            containerNumber: containerNumber,
            sealNumber: sealNumber,
            netWeight: netWeight,
            tareWeight: tareWeight,
            grossWeight: grossWeight
        });
    });
    if (hasError) {
        setPackingListMessage('Enter valid numbers for net/tare/gross weight.', true);
        return;
    }
    packingListEntries = normalizePackingListEntries(rows, containerCount);
    const payload = buildPackingListPayload(packingListEntries);
    setPackingListMessage('Saving packing list...', false);
    $('.packing-list-save-btn').prop('disabled', true);
    JSUTIL.callAJAXPost(
        JSUTIL.getUrl('/data/ProductContainer/upsert_multiple'),
        JSON.stringify(payload),
        function (res) {
            const list = Array.isArray(res?.data) ? res.data : (Array.isArray(res) ? res : []);
            if (Array.isArray(list) && list.length) {
                packingListEntries = normalizePackingListEntries(list, containerCount);
            }
            persistPackingListEntries();
            setPackingListEditMode(false);
            setPackingListMessage('Packing list saved.', false);
            $('.packing-list-save-btn').prop('disabled', false);
        },
        function (err) {
            console.error('Failed to save packing list', err);
            setPackingListMessage('Could not save packing list. Please try again.', true);
            $('.packing-list-save-btn').prop('disabled', false);
        }
    );
}

function handlePrintFinalLabel() {
    if (!currentBPCRId) {
        return;
    }
    const result = buildFinalProductPrintData();
    if (!result.valid) {
        const errorMsg = result.message || 'Final product details are missing.';
        if (typeof JSUTIL !== 'undefined' && typeof JSUTIL.buildErrorModal === 'function') {
            JSUTIL.buildErrorModal(errorMsg);
        } else {
            JSUTIL.buildErrorModal(errorMsg);
        }
        return;
    }
    openBpcrLabelModal(result.data);
}

function handlePrintReport() {
    if (!currentBPCRId) {
        return;
    }
    const result = buildFinalProductPrintData();
    if (!result.valid) {
        const message = result.message || 'Finalize the product details before printing the report.';
        if (typeof JSUTIL !== 'undefined' && typeof JSUTIL.buildErrorModal === 'function') {
            JSUTIL.buildErrorModal(message);
        } else {
            JSUTIL.buildErrorModal(message);
        }
        return;
    }
    const groups = groupBPCRMaterials(bpcrMaterialItems || []);
    const reportHtml = buildBPCRPrintReportHtml(groups, result.data);
    openBPCRReportPrintWindow(reportHtml, result.data);
}

function buildBPCRPrintReportHtml(groups, finalData) {
    const pages = [];
    pages.push(buildMaterialPrintPage('Raw Material Report', 'Raw materials planned for this BPCR.', groups.raw));
    pages.push(buildMaterialPrintPage('Packing Material Report', 'Packing materials planned for this BPCR.', groups.packing));
    pages.push(buildMaterialPrintPage('General Material Report', 'General materials planned for this BPCR.', groups.general));
    pages.push(buildFinalProductReportPage(finalData));
    return pages.join('');
}

function buildMaterialPrintPage(title, description, items, showIssued) {
    const columnCount = 8;
    const rows = (items && items.length)
        ? items.map(function (item) {
            return buildMaterialPrintRow(item);
        }).join('')
        : `<tr><td colspan="${columnCount}" class="empty-cell">No materials available.</td></tr>`;
    return `
        <div class="print-page">
            <h2>${formatPrintValue(title)}</h2>
            <p class="page-note">${formatPrintValue(description)}</p>
            <table>
                <thead>
                    <tr>
                        <th>Material Name</th>
                        <th>Material Code</th>
                        <th>UOM</th>
                        <th>Standard Qty</th>
                        <th>Requested Qty</th>
                        <th>Issued Qty</th>
                        <th>Used Qty</th>
                        <th>Returned Qty</th>
                    </tr>
                </thead>
                <tbody>
                    ${rows}
                </tbody>
            </table>
        </div>
    `;
}

function buildMaterialPrintRow(item) {
    const product = item?.product || {};
    const plannedQty = firstDefinedNumber(item?.plannedQuantity, 0);
    const requestedQty = firstDefinedNumber(item?.actualQuantity, 0);
    const issuedQty = getIssuedQuantityForProduct(product?.id, item?.productLot || item?.ProductLot || '');
    const usedQty = firstDefinedNumber(item?.quantityConsumed, item?.quantityUsed, 0);
    const returnedQty = issuedQty - usedQty;
    return `
        <tr>
            <td>${formatPrintValue(product?.name || '')}</td>
            <td>${formatPrintValue(product?.number || '')}</td>
            <td>${formatPrintValue(product?.uom || '')}</td>
            <td>${formatPrintValue(plannedQty)}</td>
            <td>${formatPrintValue(requestedQty)}</td>
            <td>${formatPrintValue(issuedQty)}</td>
            <td>${formatPrintValue(usedQty)}</td>
            <td>${formatPrintValue(returnedQty)}</td>
        </tr>
    `;
}

function shouldShowIssuedQtyInPrint() {
    if (!bpcrIssuedQuantityMap) {
        return false;
    }
    return Object.keys(bpcrIssuedQuantityMap).some(function (key) {
        const entry = bpcrIssuedQuantityMap[key];
        if (typeof entry === 'number') {
            return Number(entry || 0) > 0;
        }
        if (entry && typeof entry === 'object') {
            if (Number(entry.total || 0) > 0) {
                return true;
            }
            return Object.keys(entry.lots || {}).some(function (lotKey) {
                return Number(entry.lots[lotKey] || 0) > 0;
            });
        }
        return false;
    });
}

function buildFinalProductReportPage(finalData) {
    const rows = [
        { label: 'Product', value: finalData.productName },
        { label: 'Product Code', value: finalData.productCode },
        { label: 'UOM', value: finalData.uom },
        { label: 'Quantity Produced', value: finalData.quantity },
        { label: 'Number of Containers / Bags', value: finalData.numberOfContainers },
        { label: 'Manufacturing Date', value: finalData.manufacturingDate },
        { label: 'Expiry / Retest Date', value: finalData.expiryDate },
        { label: 'Date Type', value: finalData.expiryType || '-' },
        { label: 'Status', value: finalData.status },
        { label: 'Doc Number', value: finalData.docNumber }
    ];
    const rowsHtml = rows.map(function (row) {
        return `
            <tr>
                <th>${formatPrintValue(row.label)}</th>
                <td>${formatPrintValue(row.value)}</td>
            </tr>
        `;
    }).join('');
    return `
        <div class="print-page">
            <h2>Final Product Summary</h2>
            <table class="final-product-table">
                <tbody>
                    ${rowsHtml}
                </tbody>
            </table>
        </div>
    `;
}

function openBPCRReportPrintWindow(contentHtml, finalData) {
    // Always open in NEW TAB
    const printWindow = window.open('', '_blank');

    if (!printWindow) {
        const message = 'Unable to open print window. Please allow pop-ups and try again.';
        if (typeof JSUTIL !== 'undefined' && typeof JSUTIL.buildErrorModal === 'function') {
            JSUTIL.buildErrorModal(message);
        } else {
            JSUTIL.buildErrorModal(message);
        }
        return;
    }

    const documentTitle = `BPCR ${formatPrintValue(finalData?.bpcrNumber || currentBPCRId)} - Production Report`;
    const html = `
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8" />
            <title>${documentTitle}</title>
            <style>
                body { font-family: "Helvetica Neue", Helvetica, Arial, sans-serif; margin: 24px; color: #111; }
                h1 { font-size: 20px; margin-bottom: 10px; }
                h2 { font-size: 16px; margin-bottom: 8px; }
                table { width: 100%; border-collapse: collapse; margin-top: 10px; }
                th, td { border: 1px solid #555; padding: 8px 10px; font-size: 13px; text-align: left; }
                th { background-color: #f3f3f3; }
                .print-page { page-break-after: always; margin-bottom: 30px; }
                .print-page:last-child { page-break-after: auto; }
                .page-note { margin: 0; color: #666; font-size: 12px; }
                .empty-cell { text-align: center; font-style: italic; color: #666; }
                .page-header { display: flex; flex-wrap: wrap; font-size: 13px; margin-bottom: 15px; }
                .page-header div { margin-right: 20px; margin-bottom: 4px; }
                .final-product-table th { width: 40%; }
            </style>
        </head>
        <body>
            <h1>BPCR Production Report</h1>
            <div class="page-header">
                <div><strong>BPCR:</strong> ${formatPrintValue(finalData.bpcrNumber)}</div>
                <div><strong>Doc Number:</strong> ${formatPrintValue(finalData.docNumber)}</div>
                <div><strong>Status:</strong> ${formatPrintValue(finalData.status)}</div>
                <div><strong>Printed On:</strong> ${formatPrintValue(finalData.printedOn)}</div>
            </div>

            ${contentHtml}

            <script>
                window.onload = function () {
                    window.focus();
                    window.print();
                };
            <\/script>
        </body>
        </html>
    `;

    printWindow.document.open();
    printWindow.document.write(html);
    printWindow.document.close();
}

function buildFinalProductPrintData() {
    const baseItem = finalProductItem || {};
    const product = baseItem.product
        || currentBPCRRecord?.productRecipe?.productStage?.product
        || currentBPCRRecord?.productRecipe?.product
        || {};
    const quantity = firstDefinedNumber(baseItem.plannedQuantity, $('#finalProductQty').val(), 0);
    const containerCount = firstDefinedNumber(baseItem.numberOfContainers, $('#finalProductNumberOfContainers').val(), 0);
    const manufacturingDateValue = baseItem.manufacturingDate || $('#finalProductManufacturingDate').val() || '';
    const expiryDateValue = baseItem.expiryDate || $('#finalProductExpiryDate').val() || '';
    const expiryType = normalizeExpiryType(
        baseItem.expiryType
        || $('#finalProductExpiryType').val()
        || 'Expiry'
    );
    const docNumber = (baseItem.docNumber
        || $('#finalProductDocNumber').val()
        || getBPCRDocNumber()
        || '').toString().trim();
    if (!Number.isFinite(quantity) || quantity <= 0) {
        return { valid: false, message: 'Final product quantity missing. Please update it before printing.' };
    }
    if (!Number.isFinite(containerCount) || containerCount <= 0) {
        return { valid: false, message: 'Final product container count missing. Please update it before printing.' };
    }
    if (!manufacturingDateValue) {
        return { valid: false, message: 'Manufacturing date missing for the final product.' };
    }
    if (!docNumber) {
        return { valid: false, message: 'BPCR document number missing. Please refresh and try again.' };
    }
    const bpcrNumber = currentBPCRRecord?.name
        || currentBPCRRecord?.number
        || currentBPCRId
        || '';
    const statusText = normalizeIssuedStatus(currentStatusValue || '') || '-';
    const data = {
        bpcrNumber: bpcrNumber,
        docNumber: docNumber,
        productName: product?.name || '',
        productCode: product?.number || '',
        uom: product?.uom || '',
        quantity: quantity,
        numberOfContainers: containerCount,
        manufacturingDate: formatDateInputValue(manufacturingDateValue) || manufacturingDateValue,
        expiryDate: formatDateInputValue(expiryDateValue) || expiryDateValue,
        expiryType: expiryType,
        status: statusText,
        printedOn: formatDateTimeDisplayValue(new Date().toISOString())
    };
    return { valid: true, data: data };
}

function openFinalProductPrintWindow(data) {
    const printWindow = window.open('', '_blank');
    if (!printWindow) {
        const message = 'Unable to open print window. Please allow pop-ups and try again.';
        if (typeof JSUTIL !== 'undefined' && typeof JSUTIL.buildErrorModal === 'function') {
            JSUTIL.buildErrorModal(message);
        } else {
            JSUTIL.buildErrorModal(message);
        }
        return;
    }
    const html = `
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8" />
            <title>BPCR ${formatPrintValue(data.bpcrNumber)} - Final Product</title>
            <style>
                body { font-family: "Helvetica Neue", Helvetica, Arial, sans-serif; margin: 24px; color: #111; }
                h1 { font-size: 20px; margin-bottom: 5px; }
                h2 { font-size: 16px; margin-top: 30px; }
                .meta-row { display: flex; flex-wrap: wrap; font-size: 14px; margin-bottom: 8px; }
                .meta-row div { margin-right: 24px; margin-bottom: 4px; }
                table { width: 100%; border-collapse: collapse; margin-top: 15px; }
                th, td { border: 1px solid #444; padding: 8px 10px; font-size: 14px; text-align: left; }
                th { background-color: #f5f5f5; font-weight: 600; }
                .footer-note { margin-top: 24px; font-size: 12px; color: #666; }
                @media print {
                    body { margin: 0; }
                    .footer-note { color: #444; }
                }
            </style>
        </head>
        <body>
            <h1>BPCR Final Product Summary</h1>
            <div class="meta-row">
                <div><strong>BPCR:</strong> ${formatPrintValue(data.bpcrNumber)}</div>
                <div><strong>Doc Number:</strong> ${formatPrintValue(data.docNumber)}</div>
                <div><strong>Status:</strong> ${formatPrintValue(data.status)}</div>
                <div><strong>Printed On:</strong> ${formatPrintValue(data.printedOn)}</div>
            </div>
            <div class="meta-row">
                <div><strong>Product:</strong> ${formatPrintValue(data.productName)}</div>
                <div><strong>Product Code:</strong> ${formatPrintValue(data.productCode)}</div>
                <div><strong>UOM:</strong> ${formatPrintValue(data.uom)}</div>
            </div>
            <table>
                <tbody>
                    <tr>
                        <th>Quantity Produced</th>
                        <td>${formatPrintValue(data.quantity)}</td>
                    </tr>
                    <tr>
                        <th>Number of Containers / Bags</th>
                        <td>${formatPrintValue(data.numberOfContainers)}</td>
                    </tr>
                    <tr>
                        <th>Manufacturing Date</th>
                        <td>${formatPrintValue(data.manufacturingDate)}</td>
                    </tr>
                    <tr>
                        <th>Expiry / Retest Date</th>
                        <td>${formatPrintValue(data.expiryDate)}</td>
                    </tr>
                    <tr>
                        <th>Date Type</th>
                        <td>${formatPrintValue(data.expiryType)}</td>
                    </tr>
                </tbody>
            </table>
            <div class="footer-note">
                Generated from BPCR ${formatPrintValue(data.bpcrNumber)}. Please ensure QC request references this printout.
            </div>
            <script>
                window.onload = function () {
                    window.focus();
                    window.print();
                };
            <\/script>
        </body>
        </html>
    `;
    printWindow.document.open();
    printWindow.document.write(html);
    printWindow.document.close();
}

function openBpcrLabelModal(finalData) {
    const records = buildBpcrLabelRecords(finalData);
    if (!records.length) {
        const message = 'Number of containers is required to print labels.';
        if (typeof JSUTIL !== 'undefined' && typeof JSUTIL.buildErrorModal === 'function') {
            JSUTIL.buildErrorModal(message);
        } else {
            JSUTIL.buildErrorModal(message);
        }
        return;
    }
    bpcrLabelState.records = records;
    bpcrLabelState.issueDate = finalData.manufacturingDate || new Date().toISOString();
    bpcrLabelState.finalData = finalData;
    populateBpcrLabelTable(records);
    setBpcrLabelModalInfo(finalData);
    setBpcrLabelModalMessage('', false);
    loadBpcrLabelPrinterOptions();
    $('#bpcrLabelModal').modal('show');
}

function buildBpcrLabelRecords(finalData) {
    const containerCount = Number(finalData.numberOfContainers) || 0;
    if (!containerCount) {
        return [];
    }
    const totalQuantity = Number(finalData.quantity) || 0;
    const perContainerQty = containerCount ? (totalQuantity / containerCount) : 0;
    const records = [];
    for (let i = 1; i <= containerCount; i += 1) {
        records.push({
            id: `bpcr-label-${i}`,
            productName: finalData.productName || '',
            productNumber: finalData.productCode || '',
            batchNumber: finalData.docNumber || finalData.bpcrNumber || '',
            uom: finalData.uom || '',
            quantity: perContainerQty,
            grossWet: perContainerQty,
            tareWet: 0,
            netWet: perContainerQty,
            containerIndex: i,
            containerTotal: containerCount,
            manufacturingDate: finalData.manufacturingDate || '',
            expiryDate: finalData.expiryDate || '',
            expiryType: finalData.expiryType || ''
        });
    }
    return records;
}

function populateBpcrLabelTable(records) {
    const body = $('#bpcrLabelTableBody');
    body.empty();
    if (!Array.isArray(records) || !records.length) {
        body.html('<tr><td colspan="6" class="text-center text-muted">No labels available.</td></tr>');
        return;
    }
    const formatInputValue = function (value) {
        return Number.isFinite(value) ? value : '';
    };
    records.forEach(function (record) {
        const grossValue = Number(record.grossWet ?? record.quantity) || 0;
        const tareValue = Number(record.tareWet) || 0;
        const netValue = Number(record.netWet ?? (grossValue - tareValue));
        body.append(`
            <tr data-record-id="${record.id}">
                <td><input type="checkbox" class="bpcr-label-row-check" checked></td>
                <td>
                    <div><strong>${record.productName || 'Product'}</strong></div>
                    <div class="text-muted small">Batch: ${record.batchNumber || '-'}</div>
                    <div class="text-muted small">Container: ${record.containerIndex} of ${record.containerTotal}</div>
                </td>
                <td>
                    <input type="number" class="form-control input-sm bpcr-label-gross-input" min="0" step="any" value="${formatInputValue(grossValue)}">
                </td>
                <td>
                    <input type="number" class="form-control input-sm bpcr-label-tare-input" min="0" step="any" value="${formatInputValue(tareValue)}">
                </td>
                <td>
                    <input type="number" class="form-control input-sm bpcr-label-net-input" min="0" step="any" value="${formatInputValue(Number.isFinite(netValue) ? netValue : Math.max(grossValue - tareValue, 0))}">
                </td>
                <td>
                    <input type="number" class="form-control input-sm bpcr-label-copy-input" min="1" value="1">
                </td>
            </tr>
        `);
    });
}

function setBpcrLabelModalInfo(finalData) {
    const parts = [];
    if (finalData.productName) {
        parts.push(`Product: ${finalData.productName}`);
    }
    if (finalData.docNumber || finalData.bpcrNumber) {
        parts.push(`Doc: ${finalData.docNumber || finalData.bpcrNumber}`);
    }
    parts.push(`Containers: ${finalData.numberOfContainers || '-'}`);
    const infoText = parts.join(' • ');
    const $info = $('#bpcrLabelInfo');
    if (infoText) {
        $info.removeClass('alert-warning').addClass('alert-info').text(infoText).show();
    } else {
        $info.hide().text('');
    }
}

function loadBpcrLabelPrinterOptions() {
    const manualPrinters = getStoredLabelPrinters();
    const options = manualPrinters.map(function (name, index) {
        return {
            id: `manual-${index}-${name}`,
            name,
            type: 'manual'
        };
    });
    const storedId = bpcrLabelState.selectedPrinterId || readStoredLabelPrinterId();
    bpcrLabelState.printerOptions = options.slice();
    populateBpcrPrinterSelect(storedId);
    if (typeof window.BrowserPrint === 'undefined') {
        return;
    }
    try {
        window.BrowserPrint.getLocalDevices(function (devices) {
            (devices || []).forEach(function (device, idx) {
                if (!device) {
                    return;
                }
                const id = `browser-${device.uid || device.name || idx}`;
                bpcrLabelState.printerOptions.push({
                    id,
                    name: device.name || device.uid || `Printer ${idx + 1}`,
                    type: 'browserprint',
                    device: device
                });
            });
            populateBpcrPrinterSelect(bpcrLabelState.selectedPrinterId || storedId);
        }, function (err) {
            console.warn('Could not load BrowserPrint printers', err);
        }, 'printer');
    } catch (err) {
        console.warn('BrowserPrint not available', err);
    }
}

function populateBpcrPrinterSelect(selectedId) {
    const $select = $('#bpcrPrinterSelect');
    $select.empty();
    if (!bpcrLabelState.printerOptions.length) {
        $select.append('<option value="">-- No printers available --</option>');
        persistLabelPrinterId('');
        return;
    }
    bpcrLabelState.printerOptions.forEach(function (option) {
        $select.append(`<option value="${option.id}">${option.name}</option>`);
    });
    let targetId = selectedId || readStoredLabelPrinterId();
    if (!bpcrLabelState.printerOptions.find(function (opt) { return opt.id === targetId; })) {
        targetId = bpcrLabelState.printerOptions[0]?.id || '';
    }
    $select.val(targetId);
    persistLabelPrinterId(targetId);
}

function readStoredLabelPrinterId() {
    if (bpcrLabelState.selectedPrinterId) {
        return bpcrLabelState.selectedPrinterId;
    }
    try {
        const stored = localStorage.getItem(BPCR_LABEL_LAST_PRINTER_KEY);
        if (stored) {
            bpcrLabelState.selectedPrinterId = stored;
        }
    } catch (err) {
        // ignore
    }
    return bpcrLabelState.selectedPrinterId || '';
}

function persistLabelPrinterId(printerId) {
    bpcrLabelState.selectedPrinterId = printerId || '';
    try {
        if (bpcrLabelState.selectedPrinterId) {
            localStorage.setItem(BPCR_LABEL_LAST_PRINTER_KEY, bpcrLabelState.selectedPrinterId);
        } else {
            localStorage.removeItem(BPCR_LABEL_LAST_PRINTER_KEY);
        }
    } catch (err) {
        // ignore
    }
}

function getStoredLabelPrinters() {
    try {
        const stored = localStorage.getItem(BPCR_LABEL_PRINTER_STORAGE_KEY);
        if (!stored) return [];
        const parsed = JSON.parse(stored);
        return Array.isArray(parsed) ? parsed : [];
    } catch (err) {
        return [];
    }
}

function saveStoredLabelPrinters(list) {
    try {
        localStorage.setItem(BPCR_LABEL_PRINTER_STORAGE_KEY, JSON.stringify(list || []));
    } catch (err) {
        console.warn('Could not persist printer list', err);
    }
}

function addBpcrManualPrinter() {
    const $input = $('#bpcrNewPrinterInput');
    const name = ($input.val() || '').toString().trim();
    if (!name) {
        return;
    }
    const printers = getStoredLabelPrinters();
    if (printers.indexOf(name) === -1) {
        printers.push(name);
        saveStoredLabelPrinters(printers);
    }
    $input.val('');
    loadBpcrLabelPrinterOptions();
    setBpcrLabelModalMessage(`Added printer "${name}".`, false);
}

function handleBpcrPrinterSelection(value) {
    persistLabelPrinterId(value || '');
    setBpcrLabelModalMessage('', false);
}

function getSelectedBpcrPrinterOption() {
    const currentId = bpcrLabelState.selectedPrinterId;
    if (!currentId) {
        return null;
    }
    return bpcrLabelState.printerOptions.find(function (opt) {
        return opt.id === currentId;
    }) || null;
}

function collectBpcrLabelSelections() {
    const selections = [];
    let invalidWeights = false;
    $('#bpcrLabelTableBody tr').each(function () {
        const $row = $(this);
        const recordId = $row.data('record-id');
        if (!recordId) {
            return;
        }
        const isChecked = $row.find('.bpcr-label-row-check').prop('checked');
        if (!isChecked) {
            return;
        }
        const gross = Number($row.find('.bpcr-label-gross-input').val());
        const tare = Number($row.find('.bpcr-label-tare-input').val());
        let net = Number($row.find('.bpcr-label-net-input').val());
        if (!Number.isFinite(gross) || gross <= 0) {
            invalidWeights = true;
            return;
        }
        const safeTare = Number.isFinite(tare) && tare >= 0 ? tare : 0;
        if (!Number.isFinite(net)) {
            net = gross - safeTare;
        }
        if (!Number.isFinite(net) || net <= 0 || net > gross) {
            invalidWeights = true;
            return;
        }
        let copies = Number($row.find('.bpcr-label-copy-input').val());
        if (!Number.isFinite(copies) || copies < 1) {
            copies = 1;
        }
        const record = bpcrLabelState.records.find(function (rec) {
            return rec.id === recordId;
        });
        if (record) {
            selections.push({
                record,
                copies,
                grossWet: gross,
                tareWet: safeTare,
                netWet: net
            });
        }
    });
    return {
        selections,
        invalidWeights
    };
}

function submitBpcrLabelPrint(forceDownload) {
    setBpcrLabelModalMessage('', false);
    const result = collectBpcrLabelSelections();
    if (result.invalidWeights) {
        setBpcrLabelModalMessage('Enter gross, tare, and net weight for each selected row.', true);
        return;
    }
    if (!result.selections.length) {
        setBpcrLabelModalMessage('Select at least one container to print.', true);
        return;
    }
    const zplContent = buildBpcrLabelZplContent(result.selections);
    if (!zplContent) {
        setBpcrLabelModalMessage('Unable to build label content.', true);
        return;
    }
    const printerOption = getSelectedBpcrPrinterOption();
    const docNumber = bpcrLabelState.finalData?.docNumber || bpcrLabelState.finalData?.bpcrNumber || 'BPCR';
    if (forceDownload || !printerOption) {
        downloadBpcrZpl(zplContent, docNumber);
        setBpcrLabelModalMessage('ZPL downloaded. Send it to your printer manually.', false);
        return;
    }
    if (printerOption.type !== 'browserprint' || !printerOption.device || typeof printerOption.device.send !== 'function') {
        downloadBpcrZpl(zplContent, docNumber);
        setBpcrLabelModalMessage('Selected printer is not available via Browser Print. ZPL downloaded instead.', true);
        return;
    }
    try {
        printerOption.device.send(zplContent, function () {
            setBpcrLabelModalMessage(`Sent labels to ${printerOption.name}.`, false);
        }, function (err) {
            console.error('BrowserPrint send error', err);
            downloadBpcrZpl(zplContent, docNumber);
            setBpcrLabelModalMessage('Printer error. ZPL downloaded instead.', true);
        });
    } catch (err) {
        console.error('BrowserPrint exception', err);
        downloadBpcrZpl(zplContent, docNumber);
        setBpcrLabelModalMessage('Could not reach the printer. ZPL downloaded instead.', true);
    }
}

function buildBpcrLabelZplContent(selections) {
    const finalData = bpcrLabelState.finalData || {};
    const baseMeta = {
        docNumber: finalData.docNumber || finalData.bpcrNumber || '',
        bpcrNumber: finalData.bpcrNumber || '',
        productName: finalData.productName || '',
        productCode: finalData.productCode || '',
        uom: finalData.uom || '',
        manufacturingDate: finalData.manufacturingDate || '',
        expiryDate: finalData.expiryDate || '',
        expiryType: finalData.expiryType || '',
        printedOn: finalData.printedOn || formatDateTimeDisplayValue(new Date().toISOString())
    };
    const chunks = [];
    selections.forEach(function (selection) {
        const record = selection.record;
        const copies = selection.copies || 1;
        const grossValue = Number(selection.grossWet) || 0;
        const tareValue = Number(selection.tareWet) || 0;
        const netValueRaw = Number(selection.netWet);
        const netValue = Number.isFinite(netValueRaw) ? netValueRaw : Math.max(grossValue - tareValue, 0);
        for (let i = 1; i <= copies; i += 1) {
            chunks.push(generateBpcrLabelZpl(record, Object.assign({}, baseMeta, {
                grossWet: grossValue,
                tareWet: tareValue,
                netWet: netValue,
                expiryType: record.expiryType || baseMeta.expiryType,
                copyIndex: i,
                totalCopies: copies
            })));
        }
    });
    return chunks.join('\n');
}

function generateBpcrLabelZpl(record, meta) {
    const copyLabel = `${record.containerIndex} of ${record.containerTotal}`;
    const uomValue = record.uom || meta.uom;
    const grossValue = Number.isFinite(meta.grossWet)
        ? meta.grossWet
        : Number(record.grossWet ?? record.quantity) || 0;
    const tareValue = Number.isFinite(meta.tareWet)
        ? meta.tareWet
        : Number(record.tareWet) || 0;
    const netValue = Number.isFinite(meta.netWet)
        ? meta.netWet
        : Math.max(grossValue - tareValue, 0);
    const grossText = formatLabelQuantity(grossValue, uomValue);
    const tareText = formatLabelQuantity(tareValue, uomValue);
    const netText = formatLabelQuantity(netValue, uomValue);
    const expiryLabelValue = normalizeExpiryType(meta.expiryType || 'Expiry');
    const sanitized = {
        title: sanitizeBpcrZplValue('UNDER TEST'),
        productName: sanitizeBpcrZplValue(record.productName || ''),
        productCode: sanitizeBpcrZplValue(record.productNumber || ''),
        docNumber: sanitizeBpcrZplValue(meta.docNumber || ''),
        batchNumber: sanitizeBpcrZplValue(record.batchNumber || ''),
        containerLabel: sanitizeBpcrZplValue(copyLabel),
        grossText: sanitizeBpcrZplValue(grossText),
        tareText: sanitizeBpcrZplValue(tareText),
        netText: sanitizeBpcrZplValue(netText),
        uom: sanitizeBpcrZplValue(uomValue || ''),
        manufacturingDate: sanitizeBpcrZplValue(formatDateInputValue(record.manufacturingDate) || record.manufacturingDate || meta.manufacturingDate || '-'),
        expiryDate: sanitizeBpcrZplValue(formatDateInputValue(record.expiryDate) || record.expiryDate || meta.expiryDate || '-'),
        expiryLabel: sanitizeBpcrZplValue(expiryLabelValue),
        printedOn: sanitizeBpcrZplValue(meta.printedOn || '')
    };
    return `
^XA
^CI28
^PW800
^LL600
^MD30
^CF0,25
^FO10,10^GB780,780,2^FS
^FO10,80^GB780,0,2^FS
^FO260,25^A0N,40,40^FD${sanitized.title}^FS
^FO20,25^A0N,28,28^FDDoc: ${sanitized.docNumber}^FS
^FO600,25^A0N,25,25^FD${sanitized.containerLabel}^FS
^FO20,105^FDProduct:^FS
^FO170,105^FD${sanitized.productName}^FS
^FO20,135^FDProduct Code:^FS
^FO170,135^FD${sanitized.productCode}^FS
^FO20,165^FDBatch No:^FS
^FO170,165^FD${currentBPCRRecord.name}^FS
^FO20,195^FDGross Wt:^FS
^FO170,195^FD${sanitized.grossText}^FS
^FO20,225^FDTare Wt:^FS
^FO170,225^FD${sanitized.tareText}^FS
^FO20,255^FDNet Wt:^FS
^FO170,255^FD${sanitized.netText}^FS
^FO20,285^FDManufacturing:^FS
^FO170,285^FD${sanitized.manufacturingDate}^FS
^FO20,315^FD${sanitized.expiryLabel}:^FS
^FO170,315^FD${sanitized.expiryDate}^FS
^FO20,375^FDPrinted On:^FS
^FO170,375^FD${sanitized.printedOn}^FS
^XZ`.trim();
}

function sanitizeBpcrZplValue(value) {
    return (value || '').toString().replace(/[\^~\\]/g, ' ').replace(/[\r\n]+/g, ' ').trim();
}

function downloadBpcrZpl(content, docNumber) {
    try {
        const blob = new Blob([content], { type: 'text/plain' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        const safeName = (docNumber || 'bpcr-labels').replace(/[^a-z0-9_-]/ig, '_');
        link.download = `${safeName}.zpl`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        setTimeout(function () {
            URL.revokeObjectURL(url);
        }, 500);
    } catch (err) {
        console.error('Failed to download ZPL', err);
    }
}

function setBpcrLabelModalMessage(message, isError) {
    const $msg = $('#bpcrLabelMessage');
    if (!message) {
        $msg.hide().text('');
        return;
    }
    $msg.removeClass('text-danger text-success').addClass(isError ? 'text-danger' : 'text-success').text(message).show();
}

function formatLabelQuantity(quantity, uom) {
    const value = Number(quantity) || 0;
    const rounded = Math.round(value * 1000) / 1000;
    const text = Number.isInteger(rounded) ? String(rounded) : rounded.toFixed(3).replace(/\.?0+$/, '');
    return `${text} ${uom || ''}`.trim();
}

function formatDateInputValue(value) {
    if (!value) {
        return '';
    }
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) {
        return '';
    }
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

function formatDateTimeDisplayValue(value) {
    if (!value) {
        return '-';
    }
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) {
        return value;
    }
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    return `${year}-${month}-${day} ${hours}:${minutes}`;
}

function formatDateTimeInputValue(value) {
    if (!value) {
        return '';
    }
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) {
        return '';
    }
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    return `${year}-${month}-${day}T${hours}:${minutes}`;
}

function parseDateTimeLocalValue(value) {
    if (!value) {
        return '';
    }
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) {
        return '';
    }
    return date.toISOString();
}

function formatPrintValue(value) {
    const displayValue = (value === 0 || value === '0')
        ? '0'
        : (value ? value : '-');
    return escapeHtml(displayValue);
}

function escapeHtml(value) {
    return String(value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function getRequestColumnLabel() {
    return isRMIssuedStatus() ? 'Requested Qty' : 'Request Qty';
}

function shouldShowRequestedColumn() {
    return true;
}

function shouldShowUsedColumn() {
    const normalized = (currentStatusValue || '').toLowerCase();
    const visibleStatuses = [
        'rm issued',
        'rm-issued',
        'product issued',
        'production output',
        'packing list creation',
        'start production',
        'in-production',
        'end production',
        'production ended',
        'complete',
        'under test',
        'under-test'
    ];
    return visibleStatuses.indexOf(normalized) !== -1 || isReleaseStatus(normalized);
}

function buildBPCRMaterialRows(tableId, list, showIssued, showRequested, showUsed, showReturned, requestLabel, typeKey) {
    let html = "";
    const materialType = getNormalizedMaterialType(typeKey);
    const items = Array.isArray(list) ? list : [];
    if (!items.length) {
        let colspan = 4;
        if (showRequested) colspan += 1;
        if (showIssued) colspan += 1;
        if (showUsed) colspan += 1;
        if (showReturned) colspan += 1;
        const label = materialType ? `${materialType} materials` : 'materials';
        html = `<tr><td colspan="${colspan}" class="text-center text-muted">No ${escapeHtml(label)}.</td></tr>`;
        $(tableId).html(html);
        setMaterialEditMode(materialEditEnabled);
        return;
    }

    items.forEach(function (item) {
        const product = item.product || {};
        const productId = product.id || '';
        const plannedQty = Number(item.plannedQuantity || 0);
        const actualQty = Number(item.actualQuantity);
        const stepName = (item.productLot || item.ProductLot || '').toString().trim();
        const normalizedStep = normalizeStepName(stepName);
        const requestedQtyValue = getRequestedQuantityForProduct(productId, stepName || normalizedStep);
        // Default to requested quantity when available; otherwise use planned qty.
        const requestQtyValue = Number.isFinite(requestedQtyValue)
            ? requestedQtyValue
            : plannedQty;
        const usedQtyValue = Number(item.quantityConsumed ?? item.quantityUsed);
        const uom = product.uom || '';
        const issuedQty = getIssuedQuantityForProduct(productId, stepName || normalizedStep);
        const parsedUsedQty = Number.isFinite(usedQtyValue) ? usedQtyValue : 0;
        const returnedQtyValue = issuedQty - parsedUsedQty;
        const requestEditable = canEditRequestedQuantities(stepName || normalizedStep);
        const issuedColumn = showIssued
            ? `<td class="issued-qty-cell">${issuedQty}</td>`
            : '';
        const usedInput = showUsed
            ? `<td class="used-qty-cell"><input type="number" class="form-control input-sm bpcr-used-qty-input" value="${Number.isFinite(usedQtyValue) ? usedQtyValue : ''}" min="0" step="any" ${Number.isFinite(issuedQty) ? `max="${issuedQty}"` : ''} ${canEditUsedQuantities() ? '' : 'readonly'}></td>`
            : '';
        const returnedColumn = showReturned
            ? `<td class="returned-qty-cell">${returnedQtyValue}</td>`
            : '';
        html += `
            <tr data-id="${item.id}" data-product-id="${productId}" data-product-type="${escapeHtml(product.type || product.Type || '')}" data-planned-qty="${plannedQty}" data-request-qty="${requestQtyValue}" data-uom="${uom}" data-material-type="${materialType}" data-step-name="${escapeHtml(normalizedStep)}" data-step-label="${escapeHtml(stepName)}">
                <td>${product.name || ''}</td>
                <td>${product.number || ''}</td>
                <td>${uom}</td>
                <td>${plannedQty}</td>
                ${showRequested
            ? `<td class="requested-qty-cell"><input type="number" class="form-control input-sm bpcr-request-qty-input" value="${requestQtyValue}" min="0" step="any" ${requestEditable ? '' : 'readonly'}></td>`
            : ''}
                ${issuedColumn}
                ${usedInput}
                ${returnedColumn}
            </tr>`;
    });

    $(tableId).html(html);
    setMaterialEditMode(materialEditEnabled);
    refreshReturnedQtyForAllRows();
}

function buildInProcessMaterialRows(tableId, list) {
    const $tbody = $(tableId);
    const $section = $('#bpcrInProcessMaterialSection');
    if (!$tbody.length) {
        return;
    }
    const items = Array.isArray(list) ? list : [];
    const hasItems = items.length > 0;
    if ($section.length) {
        $section.attr('data-has-materials', hasItems ? 'true' : 'false');
    }
    if (!hasItems) {
        $tbody.html('<tr><td colspan="4" class="text-center text-muted">No in-process materials.</td></tr>');
        if ($section.length) {
            $section.hide();
        }
        return;
    }
    if ($section.length) {
        $section.show();
    }
    const rows = items.map(function (item) {
        const product = item.product || {};
        const productId = product.id || '';
        const plannedQty = Number(item.plannedQuantity || 0);
        const uom = product.uom || '';
        const stepName = (item.productLot || item.ProductLot || '').toString().trim();
        const normalizedStep = normalizeStepName(stepName);
        return `
            <tr data-in-process="true" data-id="${item.id || ''}" data-product-id="${productId}" data-planned-qty="${plannedQty}" data-uom="${uom}" data-material-type="inprocess" data-step-name="${escapeHtml(normalizedStep)}" data-step-label="${escapeHtml(stepName)}">
                <td>${product.name || ''}</td>
                <td>${product.number || ''}</td>
                <td>${uom}</td>
                <td>${plannedQty}</td>
            </tr>
        `;
    }).join('');
    $tbody.html(rows);
}

function getIssuedQuantityForProduct(productId, lotLabel) {
    if (!productId) {
        return 0;
    }
    const entry = bpcrIssuedQuantityMap[String(productId)];
    const normalizedLot = normalizeLotName(lotLabel || '');
    if (typeof entry === 'number') {
        return Number(entry) || 0;
    }
    if (entry && typeof entry === 'object') {
        if (normalizedLot && entry.lots && Number.isFinite(entry.lots[normalizedLot])) {
            return entry.lots[normalizedLot];
        }
        if (Number.isFinite(entry.total)) {
            return entry.total;
        }
    }
    return 0;
}

function updateReturnedQuantityForRow($row) {
    if (!$row || !$row.length) {
        return;
    }
    const productId = $row.data('product-id');
    const lotLabel = ($row.data('step-name') || $row.data('step-label') || '').toString();
    const issuedQty = getIssuedQuantityForProduct(productId, lotLabel);
    const usedVal = Number($row.find('.bpcr-used-qty-input').val());
    const safeIssued = Number.isFinite(issuedQty) ? issuedQty : 0;
    const safeUsed = Number.isFinite(usedVal) ? usedVal : 0;
    const returned = Math.max(safeIssued - safeUsed, 0);
    $row.find('.returned-qty-cell').text(returned.toFixed(2));
}

function refreshReturnedQtyForAllRows() {
    $('#editBPCRPanel').find('tr').each(function () {
        updateReturnedQuantityForRow($(this));
    });
}

function getRequestedQuantityForProduct(productId, lotLabel) {
    if (!productId) {
        return null;
    }
    const entry = bpcrRequestedQuantityMap[String(productId)];
    if (entry === undefined || entry === null) {
        return null;
    }
    const normalizedLot = normalizeLotName(lotLabel || '');
    if (typeof entry === 'number') {
        return Number.isFinite(entry) ? entry : null;
    }
    if (normalizedLot && entry?.lots && Number.isFinite(entry.lots[normalizedLot])) {
        return entry.lots[normalizedLot];
    }
    if (Number.isFinite(entry?.total)) {
        return entry.total;
    }
    return null;
}

function toggleIssuedQtyHeaders(showIssued) {
    const display = showIssued ? 'table-cell' : 'none';
    $('.issued-qty-header').css('display', display);
    if (!showIssued) {
        return;
    }
}

function toggleRequestedQtyHeaders(showRequested, labelText) {
    const display = showRequested ? 'table-cell' : 'none';
    $('.requested-qty-header').css('display', display);
    $('.requested-qty-cell').css('display', display);
    if (labelText) {
        $('.requested-qty-header').text(labelText);
    }
}

function refreshRequestQtyInputsReadonly() {
    $('#bpcr-material .bpcr-request-qty-input').each(function () {
        const $input = $(this);
        const $row = $input.closest('tr');
        const rowLot = ($row.data('step-name') || $row.data('step-label') || '').toString();
        const editable = canEditRequestedQuantities(rowLot);
        $input.prop('readonly', !editable);
    });
}

function toggleUsedQtyHeaders(showUsed) {
    const display = showUsed ? 'table-cell' : 'none';
    $('.used-qty-header').css('display', display);
    $('.used-qty-cell').css('display', display);
}

function toggleReturnedQtyHeaders(showReturned) {
    const display = showReturned ? 'table-cell' : 'none';
    $('.returned-qty-header').css('display', display);
    $('.returned-qty-cell').css('display', display);
}

function getNormalizedMaterialType(value) {
    return (value || '').toString().toLowerCase().replace(/\s+/g, '');
}

function normalizeStepName(value) {
    return (value || '').toString().toLowerCase().trim();
}

function getUniqueMaterialSteps(items) {
    const seen = {};
    const steps = [];
    (items || []).forEach(function (item) {
        const label = (item && (item.productLot || item.ProductLot))
            ? (item.productLot || item.ProductLot).toString().trim()
            : '';
        const key = normalizeStepName(label);
        if (!label || !key || seen[key]) {
            return;
        }
        seen[key] = true;
        steps.push(label);
    });
    return steps;
}

function refreshMaterialTypeFilterOptions(items) {
    const $filter = $('#bpcrMaterialTypeFilter');
    if (!$filter.length) {
        return;
    }
    const previousValue = $filter.val();
    const baseOptions = MATERIAL_TYPE_FILTER_OPTIONS.map(function (opt) {
        return `<option value="${opt.value}" data-filter-type="type">${opt.label}</option>`;
    }).join('');
    const stepOptions = getUniqueMaterialSteps(items).map(function (label) {
        const normalized = normalizeStepName(label);
        const safeLabel = escapeHtml(label);
        const safeNormalized = escapeHtml(normalized);
        return `<option value="${safeLabel}" data-filter-type="step" data-step-name="${safeNormalized}">${safeLabel}</option>`;
    }).join('');
    $filter.html(baseOptions + stepOptions);
    const hasPrevious = previousValue
        && $filter.find('option').filter(function () { return $(this).val() === previousValue; }).length;
    if (hasPrevious) {
        $filter.val(previousValue);
    } else if ($filter.find('option').filter(function () { return $(this).val() === MATERIAL_TYPE_FILTER_ALL; }).length) {
        $filter.val(MATERIAL_TYPE_FILTER_ALL);
    }
    $('.select2').select2();
}

function applyMaterialTypeFilter(selectedType) {
    const $filter = $('#bpcrMaterialTypeFilter');
    const currentValue = selectedType || ($filter.length ? $filter.val() : '') || MATERIAL_TYPE_FILTER_ALL;
    const $selectedOption = $filter.length
        ? $filter.find('option').filter(function () { return $(this).val() === currentValue; }).first()
        : $();
    const filterType = ($selectedOption.data('filter-type') || '').toString();
    const stepFilter = filterType === 'step'
        ? normalizeStepName($selectedOption.data('step-name') || currentValue)
        : '';
    const normalizedSelected = $selectedOption.length
        ? getNormalizedMaterialType(currentValue)
        : MATERIAL_TYPE_FILTER_ALL;
    const showAll = !normalizedSelected || normalizedSelected === MATERIAL_TYPE_FILTER_ALL;
    const $sections = $('#bpcr-material .material-section[data-material-type]');
    $sections.find('tbody tr').show();

    if (stepFilter) {
        $sections.each(function () {
            const $section = $(this);
            const $rows = $section.find('tbody tr');
            let hasVisibleRows = false;
            $rows.each(function () {
                const rowStep = normalizeStepName($(this).data('step-name') || $(this).data('step-label') || '');
                const showRow = rowStep === stepFilter;
                $(this).toggle(showRow);
                if (showRow) {
                    hasVisibleRows = true;
                }
            });
            $section.toggle(hasVisibleRows);
        });
        toggleAssignButton(currentStatusValue);
        return;
    }

    $sections.each(function () {
        const $section = $(this);
        const sectionType = getNormalizedMaterialType($section.data('material-type'));
        const hasMaterialsAttr = $section.attr('data-has-materials');
        const hasMaterials = typeof hasMaterialsAttr === 'undefined'
            ? true
            : String(hasMaterialsAttr).toLowerCase() === 'true';
        const shouldShow = hasMaterials && (showAll || sectionType === normalizedSelected);
        $section.toggle(shouldShow);
    });
    toggleAssignButton(currentStatusValue);
}

function fetchIssuedMaterialQuantities(bpcrIdParam, callback) {
    const bpcrId = bpcrIdParam || currentBPCRId;
    if (!bpcrId) {
        resetRequestedMaterialLots();
        materialRequestCheckPending = false;
        toggleAssignButton(currentStatusValue);
        if (typeof callback === 'function') {
            callback({}, []);
        }
        return;
    }
    materialRequestCheckPending = true;
    toggleAssignButton(currentStatusValue);
    const remarkPattern = `(ID:${bpcrId})`;
    const requestPayload = {
        fields: 'Id;Remarks;Number;ProcessStatus',
        conditions: [`Remarks LIKE '%${remarkPattern}%'`],
        logic: '{0}'
    };
    JSUTIL.callAJAXPost('/data/ProductRequest/list',
        JSON.stringify(requestPayload),
        function (res) {
            const list = Array.isArray(res?.data) ? res.data : (res || []);
            const requestIds = list.map(function (item) { return item?.id || item?.Id; }).filter(Boolean);
        const requestMeta = {};
        hasExistingMaterialRequest = requestIds.length > 0;
        materialRequestCheckPending = false;
        resetRequestedMaterialLots();
        list.forEach(function (r) {
            const rid = r?.id || r?.Id;
            if (!rid) { return; }
            const remarks = r?.remarks || r?.Remarks || '';
            const lotFromRemarks = parseRequestedLotFromRemarks(remarks);
            const statusValue = r?.processStatus || r?.ProcessStatus || '';
            const isCancelled = isCancelledRequestStatus(statusValue);
            if (lotFromRemarks && !isCancelled) {
                markRequestedLot(lotFromRemarks);
            }
            requestMeta[String(rid)] = {
                id: rid,
                number: r?.number || r?.Number || '',
                remarks: remarks,
                lot: lotFromRemarks,
                status: statusValue
            };
        });
            if (!requestIds.length) {
                toggleAssignButton(currentStatusValue);
                if (typeof callback === 'function') {
                    callback({}, []);
                }
                return;
            }
            const idString = requestIds.map(function (id) { return `'${id}'`; }).join(',');
            const itemPayload = {
                fields: 'Id;Product.Id;Product.Name;Product.Number;QuantityIssued;QuantityRequested;ProductRequest.Id',
                conditions: [`ProductRequest.Id IN (${idString})`],
                logic: '{0}'
            };
            JSUTIL.callAJAXPost('/data/ProductRequestItem/list',
                JSON.stringify(itemPayload),
                function (itemRes) {
                    const data = Array.isArray(itemRes?.data) ? itemRes.data : (itemRes || []);
                    const issuedMap = {};
                    const requestedMap = {};
                    const details = [];
                    data.forEach(function (row) {
                        const pid = row?.product?.id || row?.product?.Id || row?.productId || '';
                        if (!pid) {
                            return;
                        }
                        const issued = Number(row?.quantityIssued || row?.QuantityIssued || 0);
                        const requested = Number(row?.quantityRequested || row?.QuantityRequested || 0);
                        const reqId = row?.productRequest?.id || row?.productRequest?.Id || row?.productRequestId || '';
                        const meta = requestMeta[String(reqId)] || {};
                        const lotLabel = meta.lot || inferLotFromProduct(pid);
                        const normalizedLot = normalizeLotName(lotLabel);
                        const productKey = String(pid);
                        const issuedEntry = issuedMap[productKey];
                        const issuedAccumulator = (issuedEntry && typeof issuedEntry === 'object')
                            ? issuedEntry
                            : { total: Number(issuedEntry || 0) || 0, lots: {} };
                        issuedAccumulator.total += issued;
                        if (normalizedLot) {
                            issuedAccumulator.lots[normalizedLot] = (issuedAccumulator.lots[normalizedLot] || 0) + issued;
                        }
                        issuedMap[productKey] = issuedAccumulator;
                        // Track only the most recent requested quantity per lot/product, not a running total.
                        const currentRequested = requestedMap[productKey] || { total: 0, lots: {} };
                        const effectiveRequested = (requested >= 0 ? requested : 0);
                        currentRequested.total = effectiveRequested;
                        if (normalizedLot) {
                            currentRequested.lots[normalizedLot] = effectiveRequested;
                        }
                        requestedMap[productKey] = currentRequested;
                        if (lotLabel && !isCancelledRequestStatus(meta.status)) {
                            markRequestedLot(lotLabel);
                        }
                        details.push({
                            requestId: reqId,
                            requestNumber: meta.number || reqId,
                            remarks: meta.remarks || '',
                            status: meta.status || '',
                            productId: pid,
                            productName: row?.product?.name || row?.product?.Name || '',
                            productCode: row?.product?.number || row?.product?.Number || '',
                            quantityRequested: requested,
                            quantityIssued: issued,
                            lotLabel: lotLabel
                        });
                    });
                    toggleAssignButton(currentStatusValue);
                    if (typeof callback === 'function') {
                        callback(issuedMap, details, requestedMap);
                    }
                },
                function (err) {
                    console.error('Failed to fetch issued material quantities', err);
                    toggleAssignButton(currentStatusValue);
                    if (typeof callback === 'function') {
                        callback({}, [], {});
                    }
                });
        },
        function (err) {
            console.error('Failed to fetch Product Requests for BPCR', err);
            hasExistingMaterialRequest = false;
            resetRequestedMaterialLots();
            materialRequestCheckPending = false;
            toggleAssignButton(currentStatusValue);
            if (typeof callback === 'function') {
                callback({}, [], {});
            }
        });
}

function normalizeIssuedStatus(status) {
    const lower = (status || '').toLowerCase();
    if (lower === 'product issued' || lower === 'rm-issued') {
        return 'RM Issued';
    }
    if (lower === 'start production') {
        return 'In-Production';
    }
    if (lower === 'end production') {
        return 'Production Ended';
    }
    return status;
}

function isRMIssuedStatus() {
    const normalized = (currentStatusValue || '').toLowerCase();
    return normalized === 'rm issued'
        || normalized === 'rm-issued'
        || normalized === 'product issued'
        || normalized === 'packing list creation'
        || normalized === 'start production'
        || normalized === 'in-production'
        || normalized === 'end production'
        || normalized === 'production ended'
        || normalized === 'complete'
        || normalized === 'under test'
        || normalized === 'under-test';
}

function shouldShowIssuedColumn() {
    const normalized = (currentStatusValue || '').toLowerCase();
    return normalized === 'rm-request'
        || normalized === 'rm-requested'
        || normalized === 'rm issued'
        || normalized === 'rm-issued'
        || normalized === 'product issued'
        || normalized === 'production output'
        || normalized === 'packing list creation'
        || normalized === 'start production'
        || normalized === 'in-production'
        || normalized === 'end production'
        || normalized === 'production ended'
        || normalized === 'complete'
        || normalized === 'under test'
        || normalized === 'under-test'
        || isReleaseStatus(normalized);
}

function isProductionEndedStatus(statusValue) {
    const normalized = (statusValue || '').toLowerCase();
    return normalized === 'production ended' || normalized === 'end production';
}

function isReleaseStatus(statusValue) {
    const normalized = (statusValue || '').toLowerCase();
    return normalized === 'release' || normalized === 'released';
}

function isNewStatus(statusValue) {
    const normalized = (statusValue || currentStatusValue || '').toLowerCase();
    return normalized === '' || normalized === 'new';
}

function isCompletedStatus(statusValue) {
    const normalized = (statusValue || '').toLowerCase();
    return normalized === 'complete'
        || normalized === 'under test'
        || normalized === 'under-test'
        || isReleaseStatus(normalized);
}

function isFinalProductLockedStatus(statusValue) {
    const normalized = (statusValue || '').toLowerCase();
    return normalized === 'production output'
        || normalized === 'packing list creation'
        || isProductionEndedStatus(statusValue)
        || isCompletedStatus(statusValue);
}

function isFinalProductEntry(item) {
    const typeValue = getMaterialTypeValue(item);
    return typeValue.indexOf('final') !== -1 || typeValue === 'in-warehouse' || typeValue === 'under-test';
}

function isBsrOrderMaterial(item) {
    const lotValue = (item?.productLot || item?.ProductLot || '').toString().trim().toLowerCase();
    return lotValue === BSR_PRODUCT_LOT.toLowerCase();
}

function selectFinalProductItem(list) {
    if (!Array.isArray(list) || !list.length) {
        return null;
    }
    const nonBsrItems = list.filter(function (item) { return !isBsrOrderMaterial(item); });
    const candidates = nonBsrItems.length ? nonBsrItems : list.slice();
    const scored = candidates.map(function (item) {
        const typeValue = getMaterialTypeValue(item);
        let score = 4;
        if (typeValue.indexOf('final') !== -1) {
            score = 1;
        } else if (typeValue.indexOf('under-test') !== -1 || typeValue === 'under-test') {
            score = 2;
        } else if (typeValue === 'in-warehouse') {
            score = 3;
        }
        return { item: item, score: score };
    }).sort(function (a, b) {
        return a.score - b.score;
    });
    return scored[0].item || null;
}

function normalizeMaterialItem(item) {
    if (!item || typeof item !== 'object') {
        return {};
    }
    const normalized = { ...item };
    normalized.id = normalized.id || normalized.Id || '';
    normalized.type = normalized.type || normalized.Type || '';
    normalized.materialStatus = normalized.materialStatus || normalized.MaterialStatus || '';
    normalized.plannedQuantity = firstDefinedNumber(
        normalized.plannedQuantity,
        normalized.PlannedQuantity
    );
    normalized.actualQuantity = firstDefinedNumber(
        normalized.actualQuantity,
        normalized.ActualQuantity
    );
    normalized.quantityConsumed = firstDefinedNumber(
        normalized.quantityConsumed,
        normalized.quantityUsed,
        normalized.QuantityConsumed,
        normalized.QuantityUsed
    );
    normalized.quantityUsed = normalized.quantityConsumed;
    normalized.actualQuantity = firstDefinedNumber(
        normalized.actualQuantity,
        normalized.ActualQuantity
    );
    normalized.numberOfContainers = firstDefinedNumber(
        normalized.numberOfContainers,
        normalized.NumberOfContainers
    );
    normalized.docNumber = normalized.docNumber || normalized.DocNumber || '';
    normalized.manufacturingDate = normalized.manufacturingDate || normalized.ManufacturingDate || '';
    normalized.expiryDate = normalized.expiryDate || normalized.ExpiryDate || '';
    normalized.expiryType = normalized.expiryType || normalized.ExpiryType || '';
    normalized.retestDate = normalized.retestDate || normalized.RetestDate || '';
    normalized.remarks = normalized.remarks || normalized.Remarks || '';
    normalized.productLot = normalized.productLot || normalized.ProductLot || '';
    normalized.product = normalizeMaterialProduct(normalized.product || normalized.Product);
    normalized.variation = firstDefinedNumber(
        normalized.variation,
        normalized.Variation
    );
    normalized.grade = normalized.grade || normalized.Grade || '';
    return normalized;
}

function normalizeMaterialProduct(product) {
    if (!product || typeof product !== 'object') {
        return {};
    }
    const normalizedProduct = { ...product };
    normalizedProduct.id = normalizedProduct.id || normalizedProduct.Id || '';
    normalizedProduct.name = normalizedProduct.name || normalizedProduct.Name || '';
    normalizedProduct.number = normalizedProduct.number || normalizedProduct.Number || '';
    normalizedProduct.uom = normalizedProduct.uom || normalizedProduct.Uom || '';
    return normalizedProduct;
}

function firstDefinedNumber() {
    for (let i = 0; i < arguments.length; i++) {
        const value = arguments[i];
        if (value === 0 || value === '0') {
            return 0;
        }
        if (value !== undefined && value !== null && value !== '') {
            const parsed = Number(value);
            if (!Number.isNaN(parsed)) {
                return parsed;
            }
        }
    }
    return 0;
}

function buildWarehouseProductCodeValue(item) {
    const baseCode = (item?.productCode
        || item?.product?.number
        || item?.product?.Number
        || '').toString().trim();
    const bpcrName = getBPCRNameForProductCode();
    if (!bpcrName) {
        return baseCode;
    }
    if (!baseCode) {
        return bpcrName;
    }
    const lowerBase = baseCode.toLowerCase();
    const lowerName = bpcrName.toLowerCase();
    if (lowerBase.indexOf(lowerName) !== -1) {
        return baseCode;
    }
    return `${baseCode} - ${bpcrName}`;
}

function getBPCRNameForProductCode() {
    const bpcrName = currentBPCRRecord?.name
        || currentBPCRRecord?.Name
        || '';
    const bpcrNumber = currentBPCRRecord?.number
        || currentBPCRRecord?.Number
        || '';
    const value = (bpcrName || bpcrNumber || currentBPCRId || '').toString().trim();
    return value;
}

function getBPCRDocNumber() {
    return (currentBPCRRecord?.number
        || currentBPCRRecord?.Number
        || currentBPCRId
        || '').toString().trim();
}

function getMaterialTypeValue(item) {
    return (item?.type || item?.Type || '').toString().toLowerCase();
}

function getBPCRPlantName(record) {
    const plant = record?.plant
        || record?.productRecipe?.plant
        || record?.productRecipe?.productStage?.plant
        || record?.productStage?.plant;
    return plant?.name || plant?.Name || '';
}

function getBPCRPlantId(record) {
    const plant = record?.plant
        || record?.productRecipe?.plant
        || record?.productRecipe?.productStage?.plant
        || record?.productStage?.plant;
    return plant?.id || plant?.Id || '';
}

function getBPCRBatchUom(record) {
    const product = record?.productRecipe?.productStage?.product
        || record?.productRecipe?.product
        || record?.productStage?.product
        || record?.product
        || {};
    return product?.uom || product?.Uom || '';
}

function getBPCRProductInfo(record) {
    const product = record?.productRecipe?.productStage?.product
        || record?.productRecipe?.product
        || record?.productStage?.product
        || record?.product
        || {};
    const number = product?.number || product?.Number || '';
    const name = product?.name || product?.Name || '';
    return {
        number: number,
        name: name,
        display: [number, name].filter(Boolean).join('  |  ')
    };
}

function getBPCRStageInfo(record) {
    const stage = record?.productRecipe?.productStage
        || record?.productStage
        || {};
    const number = stage?.number || stage?.Number || stage?.code || stage?.Code || '';
    const name = stage?.name || stage?.Name || '';
    return {
        number: number,
        name: name,
        display: [number, name].filter(Boolean).join('  |  ')
    };
}

function setCurrentStageFinalFlag(record) {
    const stage = getProductStageObject(record);
    const explicitFlag = getStageFinalFlag(stage);
    if (explicitFlag !== null) {
        currentStageIsFinal = explicitFlag;
        togglePackingListStep(explicitFlag);
        toggleQualityControlStep(explicitFlag);
        updateTransferTypeOptions();
        toggleProductTransferButtons(currentStatusValue);
        return;
    }
    currentStageIsFinal = false;
    togglePackingListStep(false);
    toggleQualityControlStep(false);
    updateTransferTypeOptions();
    toggleProductTransferButtons(currentStatusValue);
    const stageId = getStageIdentifier(stage, record);
    if (!stageId) {
        return;
    }
    fetchStageFinalFlag(stageId, function (isFinal) {
        currentStageIsFinal = isFinal;
        togglePackingListStep(isFinal);
        toggleQualityControlStep(isFinal);
        toggleQCButton(currentStatusValue);
        updateTransferTypeOptions();
        toggleProductTransferButtons(currentStatusValue);
        renderFinalProductSection(finalProductItem);
    });
}

function getProductStageObject(record) {
    const source = record || currentBPCRRecord || {};
    return source?.productStage
        || source?.ProductStage
        || source?.productRecipe?.productStage
        || source?.productRecipe?.ProductStage
        || null;
}

function getStageFinalFlag(stage) {
    if (!stage) {
        return null;
    }
    if (typeof stage.isFinalStage === 'boolean') {
        return stage.isFinalStage;
    }
    if (typeof stage.IsFinalStage === 'boolean') {
        return stage.IsFinalStage;
    }
    if (typeof stage.isFinalStage === 'string') {
        return stage.isFinalStage.toLowerCase() === 'yes'
            || stage.isFinalStage.toLowerCase() === 'true';
    }
    if (typeof stage.IsFinalStage === 'string') {
        return stage.IsFinalStage.toLowerCase() === 'yes'
            || stage.IsFinalStage.toLowerCase() === 'true';
    }
    if (stage.isFinalStage === 1 || stage.IsFinalStage === 1) {
        return true;
    }
    return null;
}

function getStageIdentifier(stage, record) {
    if (stage && (stage.id || stage.Id)) {
        return stage.id || stage.Id;
    }
    const source = record || currentBPCRRecord || {};
    return source?.productStageId || source?.ProductStageId || '';
}

function fetchStageFinalFlag(stageId, onComplete) {
    if (!stageId) {
        if (typeof onComplete === 'function') {
            onComplete(false);
        }
        return;
    }
    const numericStageId = Number(stageId);
    const sanitizedStageId = String(stageId).replace(/'/g, "''");
    const condition = Number.isNaN(numericStageId)
        ? `Id = '${sanitizedStageId}'`
        : `Id = ${numericStageId}`;
    const payload = {
        fields: 'Id;IsFinalStage;isFinalStage',
        conditions: [condition],
        logic: '{0}'
    };
    JSUTIL.callAJAXPost('/data/ProductStage/list',
        JSON.stringify(payload),
        function (res) {
            const list = Array.isArray(res?.data)
                ? res.data
                : (Array.isArray(res) ? res : []);
            const stage = list[0] || {};
            const isFinal = stage.isFinalStage === true || stage.IsFinalStage === true;
            if (typeof onComplete === 'function') {
                onComplete(isFinal);
            }
        },
        function (err) {
            console.error('Failed to fetch product stage details', err);
            if (typeof onComplete === 'function') {
                onComplete(false);
            }
        });
}

function toggleQualityControlStep(show) {
    $('.step-button[data-step="6"]').toggle(!!show);
}

function togglePackingListStep(show) {
    $('.step-button[data-step="5"]').toggle(!!show);
}

function isFinalStageForCurrentBatch() {
    return currentStageIsFinal === true;
}

function loadBOMMetaForHeader(record) {
    currentBOMMeta = null;
    const bomId = getBPCRBomId(record);
    currentBOMId = bomId;
    if (!bomId) {
        return;
    }
    setBOMMetaDisplay({});
    fetchBOMMeta(bomId,
        function (meta) {
            if (currentBOMId !== bomId || !meta) {
                return;
            }
            currentBOMMeta = meta;
            setBOMMetaDisplay({
                batchSize: meta.batchSize,
                variation: meta.variation,
                uom: meta.uom || getBPCRBatchUom(record)
            });
            const bomLabel = meta.masterBatchRecord || meta.name || '';
            const currentLabel = ($('#bpcrBOMDisplay').text() || '').trim();
            if (bomLabel && (!currentLabel || currentLabel === '-')) {
                $('#bpcrBOMDisplay').text(bomLabel);
            }
        },
        function (err) {
            console.error('Failed to load BOM details for batch header', err);
        });
}

function setBatchMetaFromOrder(record) {
    const batchSize = getNullableNumber(record?.quantity, record?.Quantity);
    const variation = getNullableNumber(record?.quantityVariation, record?.QuantityVariation);
    const uom = getBPCRBatchUom(record);
    setBOMMetaDisplay({
        batchSize: batchSize,
        variation: variation,
        uom: uom
    });
}

function setBOMMetaDisplay(meta) {
    const safeMeta = meta || {};
    const mergedMeta = {
        ...safeMeta,
        uom: safeMeta.uom || getBPCRBatchUom(currentBPCRRecord)
    };
    $('#bpcrBatchSizeDisplay').text(formatBatchSizeDisplay(mergedMeta));
}

function getBOMProductUom(item) {
    return item?.product?.uom
        || item?.product?.Uom
        || item?.productStage?.product?.uom
        || item?.productStage?.product?.Uom
        || item?.ProductStage?.Product?.Uom
        || item?.Product?.Uom
        || '';
}

function fetchBOMMeta(bomId, onSuccess, onError) {
    const condition = /^[0-9]+$/.test(String(bomId))
        ? 'Id = ' + bomId
        : "Id = '" + bomId + "'";
    const payload = {
        fields: 'Id;MasterBatchRecord;Name;ProductQuantity;QuantityVariation;Product.Uom;ProductStage.Product.Uom',
        conditions: [condition],
        logic: '{0}'
    };
    JSUTIL.callAJAXPost('/data/ProductRecipe/list',
        JSON.stringify(payload),
        function (res) {
            const list = Array.isArray(res?.data)
                ? res.data
                : (Array.isArray(res) ? res : []);
            const item = list[0] || null;
            if (!item) {
                if (typeof onSuccess === 'function') {
                    onSuccess(null);
                }
                return;
            }
            let batchSize = item.productQuantity;
            if (batchSize === undefined || batchSize === null || batchSize === '') {
                batchSize = item.ProductQuantity;
            }
            let variation = item.quantityVariation;
            if (variation === undefined || variation === null || variation === '') {
                variation = item.QuantityVariation;
            }
            const uom = getBOMProductUom(item) || getBPCRBatchUom(currentBPCRRecord);
            const meta = {
                id: item.id || item.Id || '',
                masterBatchRecord: item.masterBatchRecord || item.MasterBatchRecord || '',
                name: item.name || item.Name || '',
                batchSize: getNullableNumber(batchSize),
                variation: getNullableNumber(variation),
                uom: uom
            };
            if (typeof onSuccess === 'function') {
                onSuccess(meta);
            }
        },
        function (err) {
            if (typeof onError === 'function') {
                onError(err);
            }
        });
}

function getBPCRBomId(record) {
    if (!record) {
        return '';
    }
    const bomId = record.productRecipe?.id
        || record.productRecipe?.Id
        || record?.ProductRecipe?.id
        || record?.ProductRecipe?.Id
        || record.totalCost
        || record.TotalCost
        || '';
    return bomId ? bomId.toString().trim() : '';
}

function getNullableNumber() {
    for (let i = 0; i < arguments.length; i++) {
        const value = arguments[i];
        if (value === undefined || value === null || value === '') {
            continue;
        }
        const parsed = Number(value);
        if (!Number.isNaN(parsed)) {
            return parsed;
        }
    }
    return null;
}

function isBatchMetaValueEmpty(value) {
    return value === undefined || value === null || value === '';
}

function formatBatchSizeDisplay(meta) {
    const batchSize = meta?.batchSize;
    const variation = meta?.variation;
    const uom = meta?.uom || '';
    if (isBatchMetaValueEmpty(batchSize)) {
        return '-';
    }
    const batchSizeText = formatBatchMetaValue(batchSize);
    const variationText = isBatchMetaValueEmpty(variation)
        ? ''
        : ` ± ${formatBatchMetaValue(variation)}`;
    const uomText = uom ? ` ${uom}` : '';
    return `${batchSizeText}${variationText}${uomText}`.trim();
}

function formatBatchMetaValue(value) {
    if (value === undefined || value === null || value === '') {
        return '-';
    }
    return value;
}
function formatDateDDMMYYYY(isoString) {
  const date = new Date(isoString);

  const day = String(date.getUTCDate()).padStart(2, '0');
  const month = String(date.getUTCMonth() + 1).padStart(2, '0');
  const year = date.getUTCFullYear();

  return `${day}/${month}/${year}`;
}
