let allBpcrRecords = [];
let plantOptions = [];
let employeeOptions = [];
let employeeTypeaheadList = [];
let stageOptions = [];
let productOptions = [];
let productTypeaheadList = [];
let bomOptions = [];
const SENT_TO_QA_STATUS = 'Sent to QA';
const QA_APPROVAL_LABEL = 'QA Approval';
const BATCH_SEQUENCE_PREFIX = 'BR-26-';
const BATCH_SEQUENCE_PAD_LENGTH = 4;

$(document).ready(function () {
    allBpcrRecords = Array.isArray(window.recordList) ? window.recordList.slice() : [];
    $('.newBPCRButton').show();
    $('.newBPCRButton').on('click', function () {
        resetNewBPCRForm();
        $('#newBPCRModal').modal('show');
        // $('selecct2').select2();
        $('.select2').select2();
    });

    loadPlantOptions();
    loadProductStageOptions();
    renderBPCRTable(allBpcrRecords);
    initializePlantFilter(allBpcrRecords);

    $('#plantFilter').on('change', function () {
        var selectedKey = $(this).val();
        var filteredRecords = filterRecordsByPlant(selectedKey);
        renderBPCRTable(filteredRecords);
    });

    $('#batchProductInput').on('input', function () {
        if (!$(this).val()) {
            setProductSelection('');
        }
    });

    $('#batchStageSelect').on('change', function () {
        var stageId = $(this).val();
        updateStageMetaPreview(stageId);
        var meta = getStageMeta(stageId);
        if (meta && meta.plantId) {
            setPlantSelection(meta.plantId);
        }
        if (meta && meta.productId) {
            setProductSelection(meta.productId, true);
        }
        resetBOMSelection();
    });

    $('#generateBatchBtn').on('click', function () {
        var stageId = $('#batchStageSelect').val();
        if (!stageId) {
            JSUTIL.buildErrorModal('Select a product stage before generating BOM options.');
            return;
        }
        loadBOMOptions(stageId);
    });

    $(document).on('click', '.delete-bom-row', function () {
        var $row = $(this).closest('tr');
        $row.remove();
        ensureBOMRowsFallback();
    });

    $(document).on('click', '.clone-bom-row', function () {
        var $row = $(this).closest('tr');
        cloneBOMRow($row);
    });

    $(document).on('input', '.batch-name-input', function () {
        this.value = (this.value || '').toUpperCase();
        $(this).css('border-color', '');
    });

    $(document).on('input', '.batch-start-date-input', function () {
        $(this).css('border-color', '');
    });

    initializeBatchStartDatePickers();

    $('#saveBPCRBtn').on('click', function () {
        var plantId = $('#batchPlantSelect').val() || '';
        var stageId = $('#batchStageSelect').val() || '';

        if (!stageId) {
            JSUTIL.buildErrorModal('Select a product stage.');
            return;
        }
        var selectedStageMeta = getStageMeta(stageId);
        if (!plantId && selectedStageMeta && selectedStageMeta.plantId) {
            plantId = selectedStageMeta.plantId;
        }
        if (!plantId) {
            JSUTIL.buildErrorModal('Select a plant for this batch.');
            return;
        }
        var rowsData = collectBOMRowData();
        if (!rowsData.length) {
            JSUTIL.buildErrorModal('Generate batch, keep at least one BOM, and provide batch names before saving.');
            return;
        }

        var $firstMissingName = $('#bomSelectionTableBody').find('tr[data-bom-id]').filter(function () {
            var val = ($(this).find('.batch-name-input').val() || '').trim();
            return !val;
        }).first();
        if ($firstMissingName.length) {
            var $input = $firstMissingName.find('.batch-name-input');
            $input.css('border-color', '#ed5565').focus();
            JSUTIL.buildErrorModal('Enter a batch name for each BOM before saving.');
            return;
        }

        var $firstMissingStart = $('#bomSelectionTableBody').find('tr[data-bom-id]').filter(function () {
            var val = ($(this).find('.batch-start-date-input').val() || '').trim();
            return !val;
        }).first();
        if ($firstMissingStart.length) {
            var $inputStart = $firstMissingStart.find('.batch-start-date-input');
            $inputStart.css('border-color', '#ed5565').focus();
            JSUTIL.buildErrorModal('Enter a start date (DD/MM/YYYY) for each batch before saving.');
            return;
        }

        var $firstInvalidStart = $('#bomSelectionTableBody').find('tr[data-bom-id]').filter(function () {
            var val = ($(this).find('.batch-start-date-input').val() || '').trim();
            return val && !isValidDDMMYYYYDate(val);
        }).first();
        if ($firstInvalidStart.length) {
            var $inputInvalidStart = $firstInvalidStart.find('.batch-start-date-input');
            $inputInvalidStart.css('border-color', '#ed5565').focus();
            JSUTIL.buildErrorModal('Enter start date in DD/MM/YYYY format.');
            return;
        }

        var invalidRow = rowsData.find(function (row) { return !row.batchName || !row.bomId || !row.startDate; });
        if (invalidRow) {
            JSUTIL.buildErrorModal('Each BOM row must have a batch name and start date before saving.');
            return;
        }

        var duplicateExisting = rowsData.find(function (row) { return isDuplicateBatchNumber(row.batchName); });
        if (duplicateExisting) {
            JSUTIL.buildErrorModal('Batch name "' + duplicateExisting.batchName + '" already exists. Use a unique value.');
            return;
        }

        var seenNames = {};
        var duplicateNew = rowsData.find(function (row) {
            var key = (row.batchName || '').toLowerCase();
            if (!key) {
                return false;
            }
            if (seenNames[key]) {
                return true;
            }
            seenNames[key] = true;
            return false;
        });
        if (duplicateNew) {
            JSUTIL.buildErrorModal('Duplicate batch names found. Please make each batch name unique.');
            return;
        }

        var createQueue = rowsData.map(function (row) {
            var bomValue = row.bomId ? Number(row.bomId) : null;
            if (row.bomId && Number.isNaN(bomValue)) {
                bomValue = row.bomId;
            }
            var plannedStartTime = convertDDMMYYYYToISO(row.startDate);

            var payload = {
                quantity: 1,
                quantityVariation: 0,
                orderStatus: SENT_TO_QA_STATUS,
                type: 'Batch',
                name: row.batchName,
                productStage: { id: stageId },
                plant: { id: plantId },
                totalCost: bomValue,
                plannedStartTime: plannedStartTime
            };
            if (employeeId) {
                payload.orderOwner = { id: employeeId };
            }
            else {
                JSUTIL.buildErrorModal('Employee Id not found. Please Contact Admin.');
            }
            return payload;
        });
        createBatchesSequentially(createQueue, function success() {
            $('#newBPCRModal').modal('hide');
            location.reload();
        }, function failure(errorName) {
            JSUTIL.buildErrorModal('Failed to create batch ' + (errorName || '') + '. Please try again.');
        });
    });
});

function resetNewBPCRForm() {
    setProductSelection('');
    setStageSelection('');
    updateStageMetaPreview('');
    setPlantSelection('');
    resetBOMSelection();
}

function loadPlantOptions(defaultPlantId) {
    var payload = {
        fields: 'Name'
    };
    JSUTIL.callAJAXPost('/data/Plant/list', JSON.stringify(payload), function (response) {
        plantOptions = (response || []).map(function (item) {
            var id = item.id || item.Id;
            if (!id) {
                return null;
            }
            var code = item.code || item.Code || '';
            var number = item.number || item.Number || '';
            var name = item.name || item.Name || '';
            var parts = [];
            if (code) {
                parts.push(code);
            } else if (number) {
                parts.push(number);
            }
            if (name) {
                parts.push(name);
            }
            return {
                id: id,
                label: parts.join(' - ') || name || ('Plant ' + id)
            };
        }).filter(function (opt) { return opt; });
        renderPlantOptions();
        var targetPlantId = defaultPlantId || '';
        if (targetPlantId) {
            setPlantSelection(targetPlantId);
        }
    });
}

function renderPlantOptions() {
    var $select = $('#batchPlantSelect');
    if (!$select.length) {
        return;
    }
    var options = ['<option value="">Select plant</option>'];
    (plantOptions || []).forEach(function (opt) {
        options.push('<option value="' + opt.id + '">' + opt.label + '</option>');
    });
    $select.html(options.join(''));
}

function setPlantSelection(plantId) {
    var $select = $('#batchPlantSelect');
    if (!$select.length) {
        return;
    }
    var value = plantId ? String(plantId) : '';
    if (!$select.find('option').length) {
        renderPlantOptions();
    }
    $select.val(value);
}

function buildProductOptions(stageList) {
    var seen = {};
    var products = [];
    (stageList || []).forEach(function (stage) {
        var productId = stage.productId;
        if (!productId || seen[productId]) {
            return;
        }
        seen[productId] = true;
        var productLabel = [stage.productNumber, stage.productName].filter(Boolean).join(' - ');
        products.push({
            id: productId,
            label: productLabel || stage.productName || ('Product ' + productId)
        });
    });
    products.sort(function (a, b) {
        return (a.label || '').localeCompare(b.label || '');
    });
    return products;
}

function buildProductTypeaheadList(list) {
    return (list || []).map(function (item) {
        var id = item.id || item.Id;
        var label = item.label || item.name || item.Name || '';
        if (!id) {
            return null;
        }
        return {
            id: id,
            text: label,
            search: (label + ' ' + id).toLowerCase()
        };
    }).filter(function (i) { return i; });
}

function initProductTypeahead() {
    var $input = $('#batchProductInput');
    var $hidden = $('#batchProductId');
    if (!$input.length || !$hidden.length || !productTypeaheadList.length) {
        return;
    }
    try {
        $input.typeahead('destroy');
    } catch (e) {
        // ignore
    }
    $input.typeahead({
        source: function (query, process) {
            query = (query || '').toLowerCase();
            var results = productTypeaheadList
                .filter(function (i) { return i.search.indexOf(query) !== -1; })
                .map(function (i) { return i.text; });
            process(results);
        },
        afterSelect: function (selectedText) {
            var chosen = productTypeaheadList.find(function (i) { return i.text === selectedText; });
            var chosenId = chosen ? chosen.id : '';
            setProductSelection(chosenId);
        }
    });
    $input.on('input', function () {
        var currentText = $(this).val();
        if (!currentText) {
            setProductSelection('');
            return;
        }
        var matched = productTypeaheadList.find(function (i) { return i.text === currentText; });
        if (!matched) {
            $hidden.val('');
            renderStageOptions('');
            setStageSelection('');
            updateStageMetaPreview('');
            resetBOMSelection();
        }
    });
}

function setProductSelection(productId, skipStageRender) {
    var $input = $('#batchProductInput');
    var $hidden = $('#batchProductId');
    if (!$input.length || !$hidden.length) {
        return;
    }
    var value = productId ? String(productId) : '';
    $hidden.val(value);
    if (!value) {
        $input.val('');
    } else {
        var matched = (productOptions || []).find(function (opt) { return String(opt.id) === value; });
        if (matched) {
            $input.val(matched.label || matched.name || '');
        }
    }
    if (!skipStageRender) {
        renderStageOptions(value);
        setStageSelection('');
        updateStageMetaPreview('');
        resetBOMSelection();
    }
}

function loadEmployeeOptions() {
    var payload = {
        fields: 'Name'
    };
    JSUTIL.callAJAXPost('/data/Employee/list', JSON.stringify(payload), function (response) {
        employeeOptions = response || [];
        employeeTypeaheadList = buildEmployeeTypeaheadList(employeeOptions);
        initOwnerInputs($('.owner-name-input'));
    });
}

function buildEmployeeTypeaheadList(list) {
    return (list || []).map(function (item) {
        var id = item.id || item.Id;
        var name = item.name || item.Name || '';
        if (!id) {
            return null;
        }
        return {
            id: id,
            text: name,
            search: (name + ' ' + id).toLowerCase()
        };
    }).filter(function (i) { return i; });
}

function initOwnerInputs($inputs) {
    if (!$inputs || !$inputs.length || !employeeTypeaheadList.length) {
        return;
    }
    $inputs.each(function () {
        var $input = $(this);
        var $hidden = $input.siblings('.owner-id-input');
        try {
            $input.typeahead('destroy');
        } catch (e) {
            // ignore
        }

        $input.typeahead({
            source: function (query, process) {
                query = (query || '').toLowerCase();
                var results = employeeTypeaheadList
                    .filter(function (i) { return i.search.indexOf(query) !== -1; })
                    .map(function (i) { return i.text; });
                process(results);
            },
            afterSelect: function (selectedText) {
                var chosen = employeeTypeaheadList.find(function (i) { return i.text === selectedText; });
                $hidden.val(chosen ? chosen.id : '');
            }
        });

        $input.on('input', function () {
            if (!$(this).val()) {
                $hidden.val('');
            }
        });
    });
}

function renderBPCRTable(list) {
    var rows = "";
    (list || []).forEach(function (u) {
        var statusLabel = buildOrderStatusLabel(u.orderStatus);
        var productName = u.productRecipe?.productStage?.product?.name
            || u.productStage?.product?.name
            || u.product?.name
            || '';
        var productNumber = u.productRecipe?.productStage?.product?.number
            || u.productStage?.product?.number
            || u.product?.number
            || '';
        var plantName = u.plant?.name
            || u.productRecipe?.plant?.name
            || u.productRecipe?.productStage?.plant?.name
            || u.productStage?.plant?.name
            || u.plant?.name
            || '';
        var productDisplay = [productNumber, productName].filter(function (p) { return p; }).join(' ');
        rows += `
        <tr>
            <td class="itemNum"><a href="javascript:void(0);" class="selectBPCRBtn" id="${u.id}">${u.number}</td>
            <td>${u.name}</td>
            <td>${productDisplay}</td>
            <td>${u.productStage.number}</td>
            <td>${plantName}</td>
            <td>${u.orderOwner?.name || 'Not Assigned'}</td>
            <td>${statusLabel}</td>
        </tr>
    `;
    });
    destroyBPCRDataTable();
    $('#allBPCRTable').html(rows);
    JSUTIL.reinitializeDataTable();
}

function destroyBPCRDataTable() {
    if (!window.$ || !$.fn || !$.fn.DataTable || !$.fn.DataTable.isDataTable) {
        return;
    }
    var $table = $('.dataTables');
    if ($table.length && $.fn.DataTable.isDataTable($table)) {
        $table.DataTable().destroy();
    }
}

function buildOrderStatusLabel(status) {
    var lowerStatus = (status || '').toLowerCase();

    var normalizedStatus =
        (lowerStatus === 'product issued' || lowerStatus === 'rm-issued')
            ? 'RM Issued'
            : (lowerStatus === 'start production'
                ? 'In-Production'
                : (lowerStatus === 'end production'
                    ? 'Production Ended'
                    : status));

    // Bigger label styling
    var baseClass = 'label';
    var extraStyle = '';

    if ((normalizedStatus || '').toLowerCase() === SENT_TO_QA_STATUS.toLowerCase()) {
        return `<span class="${baseClass} label-info" style="${extraStyle}">
                    ${QA_APPROVAL_LABEL}
                </span>`;
    }

    if (['under test', 'under-test'].includes((normalizedStatus || '').toLowerCase())) {
        return `<span class="${baseClass} label-warning" style="${extraStyle}">
                    ${normalizedStatus}
                </span>`;
    }

    if (normalizedStatus === 'QC-Approved') {
        return `<span class="${baseClass} label-success" style="${extraStyle}">
                    ${normalizedStatus}
                </span>`;
    }

    if (normalizedStatus === 'QC-Rejected') {
        return `<span class="${baseClass} label-danger" style="${extraStyle}">
                    ${normalizedStatus}
                </span>`;
    }

    return `<span class="${baseClass} label-default" style="${extraStyle}">
                ${normalizedStatus || ''}
            </span>`;
}

function initializePlantFilter(records) {
    var $filter = $('#plantFilter');
    if (!$filter.length) {
        return;
    }
    var options = ['<option value="">All Plants</option>'];
    var seen = {};

    (records || []).forEach(function (record) {
        var plantInfo = getRecordPlantInfo(record);
        if (!plantInfo || seen[plantInfo.key]) {
            return;
        }
        seen[plantInfo.key] = true;
        options.push(`<option value="${plantInfo.key}">${plantInfo.label}</option>`);
    });
    $filter.html(options.join(''));
}

function filterRecordsByPlant(selectedKey) {
    var key = (selectedKey || '').trim();
    if (!key) {
        return allBpcrRecords.slice();
    }
    return allBpcrRecords.filter(function (record) {
        var plantInfo = getRecordPlantInfo(record);
        if (!plantInfo || !plantInfo.matches.length) {
            return false;
        }
        return plantInfo.matches.indexOf(key) !== -1;
    });
}

function getRecordPlantInfo(record) {
    if (!record) {
        return null;
    }
    var plant = record.plant
        || record.plan
        || record.productRecipe?.plant
        || record.productRecipe?.plan
        || record.productRecipe?.productStage?.plant
        || record.productRecipe?.productStage?.plan
        || record.productStage?.plant
        || record.productStage?.plan;
    if (!plant) {
        return null;
    }
    var id = plant.id || plant.Id || '';
    var number = plant.number || plant.Number || '';
    var name = plant.name || plant.Name || '';
    var keySource = id || number || name;
    if (!keySource) {
        return null;
    }
    var key = String(keySource);
    var matches = [];
    if (key) {
        matches.push(key);
    }
    if (id) {
        var idStr = String(id);
        if (matches.indexOf(idStr) === -1) {
            matches.push(idStr);
        }
    }
    if (number) {
        var numberStr = String(number);
        if (matches.indexOf(numberStr) === -1) {
            matches.push(numberStr);
        }
    }
    if (name) {
        var nameStr = String(name);
        if (matches.indexOf(nameStr) === -1) {
            matches.push(nameStr);
        }
    }

    var labelParts = [];
    if (number) {
        labelParts.push(number);
    } else if (id) {
    }
    if (name) {
        labelParts.push(name);
    }
    var label = labelParts.join(' - ') || 'Plant';

    return {
        key: key,
        label: label,
        matches: matches
    };
}

function isDuplicateBatchNumber(batchNumber) {
    if (!batchNumber) {
        return false;
    }
    var target = batchNumber.toString().toLowerCase();
    return (allBpcrRecords || []).some(function (record) {
        var name = (record?.name || record?.Name || '').toString().toLowerCase();
        var number = (record?.number || record?.Number || '').toString().toLowerCase();
        return name === target || number === target;
    });
}

function loadProductStageOptions(defaultStageId) {
    var payload = {
        fields: 'Id;Name;Number;Product.Id;Product.Name;Product.Number;Plant.Id;Plant.Name;Plant.Number;'
    };
    JSUTIL.callAJAXPost('/data/ProductStage/list', JSON.stringify(payload), function (response) {
        stageOptions = (response || []).map(function (item) {
            var id = item.id || item.Id;
            if (!id) {
                return null;
            }
            var product = item.product || item.Product || {};
            var plant = item.plant || item.Plant || {};
            return {
                id: id,
                name: item.name || item.Name || '',
                code: item.code || item.Code || item.number || item.Number || '',
                productId: product.id || product.Id || '',
                productName: product.name || product.Name || '',
                productNumber: product.number || product.Number || '',
                plantId: plant.id || plant.Id || '',
                plantName: plant.name || plant.Name || ''
            };
        }).filter(function (opt) { return opt; });
        productOptions = buildProductOptions(stageOptions);
        productTypeaheadList = buildProductTypeaheadList(productOptions);
        initProductTypeahead();
        var selectedProductId = $('#batchProductId').val() || '';
        renderStageOptions(selectedProductId);
        if (defaultStageId) {
            setStageSelection(defaultStageId);
        }
    });
}

function renderStageOptions(filterProductId) {
    var $select = $('#batchStageSelect');
    if (!$select.length) {
        return;
    }
    var currentValue = $select.val();
    var hasCurrent = false;
    var options = ['<option value="">Select product stage</option>'];
    (stageOptions || []).forEach(function (opt) {
        if (filterProductId && String(opt.productId) !== String(filterProductId)) {
            return;
        }
        if (String(opt.id) === String(currentValue)) {
            hasCurrent = true;
        }
        var productPart = [opt.productNumber, opt.productName].filter(Boolean).join(' - ');
        var stagePart = [opt.code, opt.name].filter(Boolean).join(' - ');
        var label = [productPart, stagePart].filter(Boolean).join(' | ');
        options.push('<option value="' + opt.id + '">' + (label || opt.name || ('Stage ' + opt.id)) + '</option>');
    });
    $select.html(options.join(''));
    if (hasCurrent) {
        $select.val(currentValue);
    } else {
        $select.val('');
    }
}

function setStageSelection(stageId) {
    var $select = $('#batchStageSelect');
    if (!$select.length) {
        return;
    }
    var value = stageId ? String(stageId) : '';
    if (!$select.find('option').length) {
        var productId = $('#batchProductId').val() || '';
        renderStageOptions(productId);
    }
    $select.val(value);
    updateStageMetaPreview(value);
    var meta = getStageMeta(value);
    if (meta && meta.plantId) {
        setPlantSelection(meta.plantId);
    }
    if (meta && meta.productId) {
        setProductSelection(meta.productId, true);
    }
}

function getStageMeta(stageId) {
    if (!stageId) {
        return null;
    }
    return (stageOptions || []).find(function (opt) { return String(opt.id) === String(stageId); }) || null;
}

function updateStageMetaPreview(stageId) {
    var $meta = $('#batchStageMeta');
    if (!$meta.length) {
        return;
    }
    var meta = getStageMeta(stageId);
    if (!meta) {
        $meta.text('');
        return;
    }
    var productPart = [meta.productNumber, meta.productName].filter(Boolean).join(' - ');
    var stagePart = [meta.code, meta.name].filter(Boolean).join(' - ');
    var plantPart = meta.plantName ? ' | Plant: ' + meta.plantName : '';
    var text = [productPart, stagePart].filter(Boolean).join(' | ') + plantPart;
    $meta.text(text);
}

function resetBOMSelection() {
    bomOptions = [];
    renderBOMPlaceholder('Select a product stage and click Generate Batch to load BOMs.');
}

function renderBOMPlaceholder(message) {
    var $section = $('#bomSelectionSection');
    var $tbody = $('#bomSelectionTableBody');
    if (!$tbody.length) {
        return;
    }
    $section.show();
    $tbody.html('<tr><td colspan="8" class="text-center text-muted">' + (message || 'No BOM available.') + '</td></tr>');
}

function loadBOMOptions(stageId) {
    var parsedStageId = Number(stageId);
    var stageCondition = Number.isNaN(parsedStageId)
        ? "ProductStage.Id = '" + stageId + "'"
        : 'ProductStage.Id = ' + parsedStageId;
    bomOptions = [];
    renderBOMPlaceholder('Loading BOMs...');
    var payload = {
        fields: 'Id;Name;MasterBatchRecord;Product.Id;Product.Uom;ProductStage.Id;ProductQuantity;',
        conditions: [stageCondition],
        logic: '{0}'
    };
    JSUTIL.callAJAXPost('/data/ProductRecipe/list',
        JSON.stringify(payload),
        function (response) {
            bomOptions = response || [];
            if (!bomOptions.length) {
                renderBOMPlaceholder('No BOM found for this stage.');
                return;
            }
            renderBOMTable(bomOptions);
        },
        function () {
            renderBOMPlaceholder('Failed to load BOMs. Try again.');
        }
    );
}

function renderBOMTable(list) {
    var $section = $('#bomSelectionSection');
    var $tbody = $('#bomSelectionTableBody');
    if (!$tbody.length) {
        return;
    }
    if (!list || !list.length) {
        renderBOMPlaceholder('No BOM found for this stage.');
        return;
    }
    var rows = (list || []).map(function (item) {
        var bomId = item.id || item.Id || '';
        var productStage = item.productStage || item.ProductStage || {};
        var product = productStage.product || productStage.Product || {};
        var stage = productStage;
        var productDisplay = [product.number || product.Number || '', product.name || product.Name || ''].filter(Boolean).join(' - ');
        var stageDisplay = [stage.code || stage.Code || stage.number || stage.Number || '', stage.name || stage.Name || ''].filter(Boolean).join(' - ');
        var masterBatchRecord = item.masterBatchRecord || item.MasterBatchRecord || '';
        var batchSize = item.productQuantity;
        if (batchSize === undefined || batchSize === null) {
            batchSize = item.ProductQuantity;
        }
        var batchUom = product.uom || product.Uom || '';
        var batchSizeDisplay = (batchSize === undefined || batchSize === null || batchSize === '') ? '-' : batchSize;
        return buildBOMRow(item, {
            productDisplay: productDisplay,
            stageDisplay: stageDisplay,
            masterBatchRecord: masterBatchRecord,
            batchSizeDisplay: batchSizeDisplay,
            batchUom: batchUom
        });
    }).join('');
    $section.show();
    $tbody.html(rows);
    initializeBatchStartDatePickers($tbody);
}

function buildBOMRow(item, displayOverrides) {
    var bomId = item && (item.id || item.Id || '');
    var productStage = item && (item.productStage || item.ProductStage) || {};
    var product = productStage.product || productStage.Product || {};
    var stage = productStage;
    var productDisplay = displayOverrides && displayOverrides.productDisplay
        ? displayOverrides.productDisplay
        : [product.number || product.Number || '', product.name || product.Name || ''].filter(Boolean).join(' - ');
    var stageDisplay = displayOverrides && displayOverrides.stageDisplay
        ? displayOverrides.stageDisplay
        : [stage.code || stage.Code || stage.number || stage.Number || '', stage.name || stage.Name || ''].filter(Boolean).join(' - ');
    var masterBatchRecord = displayOverrides && displayOverrides.masterBatchRecord !== undefined
        ? displayOverrides.masterBatchRecord
        : item && (item.masterBatchRecord || item.MasterBatchRecord || '');
    var batchSize = item ? item.productQuantity : '';
    if (batchSize === undefined || batchSize === null) {
        batchSize = item ? item.ProductQuantity : '';
    }
    var batchSizeDisplay = displayOverrides && displayOverrides.batchSizeDisplay !== undefined
        ? displayOverrides.batchSizeDisplay
        : (batchSize === undefined || batchSize === null || batchSize === '') ? '-' : batchSize;
    var batchNameValue = (displayOverrides && displayOverrides.batchName !== undefined) ? displayOverrides.batchName : '';
    var startDateValue = (displayOverrides && displayOverrides.startDate !== undefined) ? displayOverrides.startDate : '';
    var batchUomValue = (displayOverrides && displayOverrides.batchUom !== undefined)
        ? displayOverrides.batchUom
        : (product.uom || product.Uom || '');
    var batchSizeWithUom = batchSizeDisplay;
    if (batchUomValue && batchSizeDisplay !== '-' && batchSizeDisplay !== '') {
        batchSizeWithUom = batchSizeDisplay + ' ' + batchUomValue;
    }

    return `
        <tr data-bom-id="${escapeHtml(bomId)}">
            <td>${escapeHtml(masterBatchRecord)}</td>
            <td>${escapeHtml(item ? (item.name || item.Name || '') : '')}</td>
            <td>${escapeHtml(productDisplay)}</td>
            <td>${escapeHtml(stageDisplay)}</td>
            <td>${escapeHtml(batchSizeWithUom)}</td>
            <td>
                <input type="text" class="form-control input-sm batch-name-input" placeholder="Enter batch name" value="${escapeHtml(batchNameValue)}">
            </td>
            <td>
                <input type="text" class="form-control datepicker input-sm batch-start-date-input" placeholder="DD/MM/YYYY" value="${escapeHtml(startDateValue)}">
            </td>
            <td class="text-center">
                <button type="button" class="btn btn-sm btn-danger delete-bom-row delete-material-row" title="Delete">
                    <i class="fa fa-trash"></i>
                </button>
                <button type="button" class="btn btn-sm btn-warning clone-bom-row" title="Clone">
                    <i class="fa fa-clone"></i>
                </button>
            </td>
        </tr>
    `;
}

function collectBOMRowData() {
    var rows = [];
    $('#bomSelectionTableBody').find('tr').each(function () {
        var $row = $(this);
        var bomId = $row.data('bom-id');
        if (!bomId) {
            return;
        }
        var batchName = ($row.find('.batch-name-input').val() || '').trim();
        var startDate = ($row.find('.batch-start-date-input').val() || '').trim();
        rows.push({
            bomId: bomId,
            batchName: batchName,
            startDate: startDate
        });
    });
    return rows;
}

function cloneBOMRow($row) {
    if (!$row || !$row.length) {
        return;
    }
    var bomId = $row.data('bom-id');
    var bomData = findBOMById(bomId);
    if (!bomData) {
        return;
    }
    var newRowHtml = buildBOMRow(bomData, {
        batchName: ''
    });
    $row.after(newRowHtml);
    initializeBatchStartDatePickers($row.next());
}

function findBOMById(bomId) {
    if (!bomId) {
        return null;
    }
    return (bomOptions || []).find(function (opt) {
        return String(opt.id || opt.Id) === String(bomId);
    }) || null;
}

function ensureBOMRowsFallback() {
    var $tbody = $('#bomSelectionTableBody');
    if (!$tbody.length) {
        return;
    }
    var hasDataRows = $tbody.find('tr[data-bom-id]').length > 0;
    if (!hasDataRows) {
        renderBOMPlaceholder('No BOM selected. Click Generate Batch to load BOMs.');
    }
}

function createBatchesSequentially(queue, onSuccess, onFailure) {
    if (!queue || !queue.length) {
        if (typeof onSuccess === 'function') {
            onSuccess();
        }
        return;
    }
    JSUTIL.callAJAXGet('/process/get_next_sequence?name=ProductionOrder',
        function (res) {
            var sharedBatchNumber = formatBatchSequenceValue(res);
            processBatchQueue(queue, sharedBatchNumber, onSuccess, onFailure);
        },
        function (err) {
            console.error('Failed to fetch ProductionOrder sequence', err);
            processBatchQueue(queue, '', onSuccess, onFailure);
        }
    );
}

function processBatchQueue(queue, sharedBatchNumber, onSuccess, onFailure) {
    if (!queue || !queue.length) {
        if (typeof onSuccess === 'function') {
            onSuccess();
        }
        return;
    }
    var remainingQueue = queue.slice();
    var current = remainingQueue.shift();
    var payload = Object.assign({}, current);
    if (!payload.name) {
        JSUTIL.buildErrorModal('Batch name missing. Please enter a batch name before saving.');
        if (typeof onFailure === 'function') {
            onFailure('');
        }
        return;
    }
    if (sharedBatchNumber) {
        payload.number = sharedBatchNumber;
    }
    JSUTIL.callAJAXPost('/data/ProductionOrder/create',
        JSON.stringify(payload),
        function () {
            processBatchQueue(remainingQueue, sharedBatchNumber, onSuccess, onFailure);
        },
        function () {
            if (typeof onFailure === 'function') {
                onFailure(payload && payload.name);
            }
        }
    );
}

function formatBatchSequenceValue(sequenceValue) {
    var raw = extractSequenceValue(sequenceValue);
    if (raw === '' || raw === null || raw === undefined) {
        return '';
    }
    var rawString = String(raw).trim();
    var matches = rawString.match(/(\d+)/g);
    var numericChunk = matches && matches.length ? matches[matches.length - 1] : rawString;
    if (!numericChunk) {
        return '';
    }
    var parsed = parseInt(numericChunk, 10);
    var base = Number.isNaN(parsed) ? numericChunk : String(parsed);
    var padded = Number.isNaN(parsed) ? base : base.padStart(BATCH_SEQUENCE_PAD_LENGTH, '0');
    return BATCH_SEQUENCE_PREFIX + padded;
}

function extractSequenceValue(sequenceResponse) {
    if (sequenceResponse === undefined || sequenceResponse === null) {
        return '';
    }
    if (typeof sequenceResponse === 'number' || typeof sequenceResponse === 'string') {
        return sequenceResponse;
    }
    if (typeof sequenceResponse === 'object') {
        var candidates = [
            sequenceResponse.sequence,
            sequenceResponse.Sequence,
            sequenceResponse.nextSequence,
            sequenceResponse.NextSequence,
            sequenceResponse.data,
            sequenceResponse.value,
            sequenceResponse.result
        ];
        for (var i = 0; i < candidates.length; i++) {
            var candidate = candidates[i];
            if (candidate !== undefined && candidate !== null) {
                return candidate;
            }
        }
        if (sequenceResponse.data && typeof sequenceResponse.data === 'object') {
            var nested = extractSequenceValue(sequenceResponse.data);
            if (nested !== '') {
                return nested;
            }
        }
        var keys = Object.keys(sequenceResponse);
        for (var j = 0; j < keys.length; j++) {
            var value = sequenceResponse[keys[j]];
            if (typeof value === 'number' || typeof value === 'string') {
                return value;
            }
        }
    }
    return '';
}

function escapeHtml(value) {
    if (value === null || value === undefined) {
        return '';
    }
    return String(value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/\"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function initializeBatchStartDatePickers(container) {
    if (!$.fn || !$.fn.datepicker) {
        return;
    }
    var $scope = container ? $(container) : $(document);
    var $inputs = $scope.find('.batch-start-date-input');
    $inputs.each(function () {
        var $input = $(this);
        if ($input.data('datepicker')) {
            $input.datepicker('destroy');
        }
        $input.datepicker({
            format: 'dd/mm/yyyy',
            autoclose: true
        });
    });
}

function parseDDMMYYYY(dateStr) {
    if (!dateStr || typeof dateStr !== 'string') {
        return null;
    }
    var parts = dateStr.split('/');
    if (parts.length !== 3) {
        return null;
    }
    var day = Number(parts[0]);
    var month = Number(parts[1]);
    var year = Number(parts[2]);
    if (!Number.isInteger(day) || !Number.isInteger(month) || !Number.isInteger(year)) {
        return null;
    }
    if (month < 1 || month > 12 || day < 1 || day > 31) {
        return null;
    }
    var utcDate = new Date(Date.UTC(year, month - 1, day));
    if (utcDate.getUTCFullYear() !== year || (utcDate.getUTCMonth() + 1) !== month || utcDate.getUTCDate() !== day) {
        return null;
    }
    return utcDate;
}

function isValidDDMMYYYYDate(dateStr) {
    return !!parseDDMMYYYY(dateStr);
}

function convertDDMMYYYYToISO(dateStr) {
    var parsed = parseDDMMYYYY(dateStr);
    return parsed ? parsed.toISOString() : '';
}
