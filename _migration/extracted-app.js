// ---------------------------------------------------------------------------
  // config.js
  // ---------------------------------------------------------------------------

  var APP_VERSION = '1.0.0';
  var CONFIG_VERSION = 1;

  var HEADER_MARKERS = {
    property: 'External ID',
    unit: 'Lease ID',
  };

  var HEADER_SEARCH_ROWS = 25;
  var GROUP_HEADER_OFFSET = 1;
  var ACCEPTED_EXTENSION = '.xlsx';

  var STANDARD_TEMPLATE_COLUMNS = [
    { id: 'property_reference', label: 'Property Reference' },
    { id: 'town', label: 'Town' },
    { id: 'address', label: 'Address' },
    { id: 'tenure_uk', label: 'Tenure (UK)' },
    { id: 'area_sqft', label: 'Area (sq ft)' },
    { id: 'unit', label: 'Unit' },
    { id: 'tenant_name', label: 'Tenant Name' },
    { id: 'start_date', label: 'Start Date' },
    { id: 'review_date', label: 'Review Date' },
    { id: 'expiry_date', label: 'Expiry Date' },
    { id: 'break_date', label: 'Break Date' },
    { id: 'mwault_year', label: 'MWAULT (Year)' },
    { id: 'contracted_rent_pa', label: 'Contracted Rent (pa)' },
    { id: 'contracted_rent_psf', label: 'Contracted Rent (psf)' },
    { id: 'deemed_rent_psf', label: 'Deemed Rent (psf)' },
    { id: 'market_rent_pa', label: 'Market Rent (pa)' },
    { id: 'market_rent_psf', label: 'Market Rent (psf)' },
    { id: 'current_void', label: 'Current Void' },
    { id: 'rent_free_month', label: 'Rent Free (Month)' },
    { id: 'expiry_void_month', label: 'Expiry Void (Month)' },
    { id: 'rent_free_expiry_month', label: 'Rent Free Expiry Month' },
    { id: 'cap_rate', label: 'Cap Rate' },
    { id: 'niy', label: 'NIY' },
    { id: 'ney', label: 'NEY' },
    { id: 'ry', label: 'RY' },
    { id: 'cv_psf', label: 'CV (psf)' },
    { id: 'cap_ex', label: 'Cap Ex' },
    { id: 'fair_value_april_2024', label: 'Fair Value April 2024' },
    { id: 'fair_value_march_2026', label: 'Fair Value March 2026' },
    { id: 'pct_change_26_april_24', label: '% Change 26 - April 24' },
    { id: 'pct_change_x_ytd', label: '% Change x YTD' },
    { id: 'comments', label: 'Comments' },
  ];

  var DEFAULT_MATCH_HINTS = {
    property: { section: 'General', field: 'Property Name' },
    unit: { section: 'General', field: 'Property Name' },
  };

  var APP_STATE = {
    property: null,
    unit: null,
    mappingConfig: null,
    mappingConfigured: false,
  };

  // ---------------------------------------------------------------------------
  // excelParser.js
  // ---------------------------------------------------------------------------

  function normaliseHeader(value) {
    if (value === null || value === undefined) return '';
    return String(value)
      .replace(/\u00a0/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase();
  }

  function normaliseMatchingKey(value) {
    return normaliseHeader(value);
  }

  function columnIndexToLetter(index) {
    var n = index + 1;
    var letter = '';
    while (n > 0) {
      var rem = (n - 1) % 26;
      letter = String.fromCharCode(65 + rem) + letter;
      n = Math.floor((n - 1) / 26);
    }
    return letter;
  }

  function buildSourceColumnKey(fileType, section, field, occurrence) {
    var parts = [
      normaliseMatchingKey(fileType),
      normaliseMatchingKey(section) || '(no section)',
      normaliseMatchingKey(field) || '(blank)',
      String(occurrence),
    ];
    return parts.join('|');
  }

  function readWorkbook(arrayBuffer) {
    return XLSX.read(arrayBuffer, {
      type: 'array',
      cellDates: true,
      dense: false,
    });
  }

  function getFirstWorksheet(workbook) {
    if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
      throw new Error('The workbook does not contain any worksheets.');
    }
    var sheetName = workbook.SheetNames[0];
    var sheet = workbook.Sheets[sheetName];
    if (!sheet) {
      throw new Error('Worksheet "' + sheetName + '" could not be read.');
    }
    return { sheetName: sheetName, sheet: sheet };
  }

  function sheetToArrayOfArrays(sheet) {
    var ref = sheet['!ref'];
    if (!ref) return [];

    var range = XLSX.utils.decode_range(ref);
    var rows = [];

    for (var r = 0; r <= range.e.r; r += 1) {
      var row = [];
      for (var c = 0; c <= range.e.c; c += 1) {
        var addr = XLSX.utils.encode_cell({ r: r, c: c });
        var cell = sheet[addr];
        if (cell === undefined || cell === null) {
          row.push(null);
        } else if (cell.v !== undefined) {
          row.push(cell.v);
        } else if (cell.w !== undefined) {
          row.push(cell.w);
        } else {
          row.push(null);
        }
      }
      rows.push(row);
    }

    return rows;
  }

  function cellDisplayText(value) {
    if (value === null || value === undefined) return '';
    if (value instanceof Date) {
      return value.toISOString();
    }
    return String(value);
  }

  function detectHeaderRow(rows, marker) {
    var target = normaliseHeader(marker);
    if (!target) {
      throw new Error('Header marker is empty.');
    }

    var limit = Math.min(rows.length, HEADER_SEARCH_ROWS);

    for (var r = 0; r < limit; r += 1) {
      var row = rows[r] || [];
      for (var c = 0; c < row.length; c += 1) {
        if (normaliseHeader(cellDisplayText(row[c])) === target) {
          return r;
        }
      }
    }

    throw new Error(
      'Could not find header marker "' + marker + '" within the first ' + HEADER_SEARCH_ROWS + ' rows.'
    );
  }

  function resolveMergedGroupHeadings(sheet, rows, groupRowIndex, columnCount) {
    var sections = Array.from({ length: columnCount }, function () {
      return '';
    });
    var groupRow = rows[groupRowIndex] || [];

    for (var c = 0; c < columnCount; c += 1) {
      var text = cellDisplayText(groupRow[c]).trim();
      if (text) sections[c] = text;
    }

    var merges = sheet['!merges'] || [];
    for (var i = 0; i < merges.length; i += 1) {
      var merge = merges[i];
      if (groupRowIndex < merge.s.r || groupRowIndex > merge.e.r) continue;

      var topLeft = rows[merge.s.r] && rows[merge.s.r][merge.s.c];
      var label = cellDisplayText(topLeft).trim();
      if (!label) continue;

      for (var mc = merge.s.c; mc <= merge.e.c; mc += 1) {
        if (mc >= 0 && mc < columnCount) {
          sections[mc] = label;
        }
      }
    }

    return sections;
  }

  function generateSourceColumns(fileType, fieldHeaders, sectionHeaders) {
    var occurrenceCounters = new Map();
    var columns = [];
    var columnCount = Math.max(fieldHeaders.length, sectionHeaders.length);

    for (var c = 0; c < columnCount; c += 1) {
      var field = (fieldHeaders[c] || '').trim();
      var section = (sectionHeaders[c] || '').trim();

      if (!field) continue;

      var occKey = normaliseMatchingKey(section) + '||' + normaliseMatchingKey(field);
      var occurrence = (occurrenceCounters.get(occKey) || 0) + 1;
      occurrenceCounters.set(occKey, occurrence);

      var columnLetter = columnIndexToLetter(c);
      var key = buildSourceColumnKey(fileType, section, field, occurrence);

      var sectionPart = section || '(No Section)';
      var occSuffix = occurrence > 1 ? ' (#' + occurrence + ')' : '';
      var displayLabel = sectionPart + ' > ' + field + occSuffix + ' [' + columnLetter + ']';

      columns.push({
        fileType: fileType,
        columnIndex: c,
        columnLetter: columnLetter,
        section: section,
        field: field,
        occurrence: occurrence,
        key: key,
        displayLabel: displayLabel,
      });
    }

    return columns;
  }

  function extractDataRows(rows, headerRowIndex, columns) {
    var parsed = [];
    var dataRowCount = 0;

    for (var r = headerRowIndex + 1; r < rows.length; r += 1) {
      var sheetRow = rows[r] || [];
      var values = {};
      var hasAnyValue = false;

      for (var i = 0; i < columns.length; i += 1) {
        var col = columns[i];
        var raw = sheetRow[col.columnIndex];
        var value = raw === undefined ? null : raw;
        values[col.key] = value;
        if (value !== null && value !== '' && !(typeof value === 'string' && value.trim() === '')) {
          hasAnyValue = true;
        }
      }

      if (!hasAnyValue) continue;

      parsed.push({
        rowIndex: r,
        values: values,
      });
      dataRowCount += 1;
    }

    return { rows: parsed, dataRowCount: dataRowCount };
  }

  function validateXlsxFile(file) {
    if (!file) {
      throw new Error('No file selected.');
    }
    var name = file.name || '';
    var lower = name.toLowerCase();
    if (!lower.endsWith('.xlsx')) {
      throw new Error('Please select an Excel .xlsx file.');
    }
  }

  function readFileAsArrayBuffer(file) {
    if (typeof file.arrayBuffer === 'function') {
      return file.arrayBuffer().catch(function () {
        return readFileAsArrayBufferFallback(file);
      });
    }
    return readFileAsArrayBufferFallback(file);
  }

  function readFileAsArrayBufferFallback(file) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onload = function () {
        resolve(reader.result);
      };
      reader.onerror = function () {
        reject(new Error('Could not read the selected file.'));
      };
      reader.readAsArrayBuffer(file);
    });
  }

  function parseSourceFile(file, fileType) {
    validateXlsxFile(file);

    var marker = HEADER_MARKERS[fileType];
    if (!marker) {
      throw new Error('Unsupported file type: ' + fileType);
    }

    return readFileAsArrayBuffer(file)
      .then(function (arrayBuffer) {
        var workbook;
        try {
          workbook = readWorkbook(arrayBuffer);
        } catch (err) {
          throw new Error('The file could not be opened as a valid Excel workbook.');
        }

        var wsInfo = getFirstWorksheet(workbook);
        var sheetName = wsInfo.sheetName;
        var sheet = wsInfo.sheet;
        var matrix = sheetToArrayOfArrays(sheet);

        if (matrix.length === 0) {
          throw new Error('The worksheet appears to be empty.');
        }

        var headerRowIndex = detectHeaderRow(matrix, marker);
        var groupRowIndex = headerRowIndex - GROUP_HEADER_OFFSET;

        var fieldRow = matrix[headerRowIndex] || [];
        var columnCount = fieldRow.length;
        var ref = sheet['!ref'];
        if (ref) {
          var range = XLSX.utils.decode_range(ref);
          columnCount = Math.max(columnCount, range.e.c + 1);
        }

        var fieldHeaders = Array.from({ length: columnCount }, function (_, c) {
          return cellDisplayText(fieldRow[c]).trim();
        });

        var sectionHeaders;
        if (groupRowIndex >= 0) {
          sectionHeaders = resolveMergedGroupHeadings(sheet, matrix, groupRowIndex, columnCount);
        } else {
          sectionHeaders = Array.from({ length: columnCount }, function () {
            return '';
          });
        }

        var columns = generateSourceColumns(fileType, fieldHeaders, sectionHeaders);

        if (columns.length === 0) {
          throw new Error('No source columns were detected on the header row.');
        }

        var extracted = extractDataRows(matrix, headerRowIndex, columns);

        return {
          fileName: file.name,
          fileType: fileType,
          sheetName: sheetName,
          headerRowIndex: headerRowIndex,
          groupRowIndex: groupRowIndex >= 0 ? groupRowIndex : null,
          columns: columns,
          rows: extracted.rows,
          dataRowCount: extracted.dataRowCount,
        };
      })
      .catch(function (err) {
        if (err && typeof err.message === 'string' && err.message) {
          throw err;
        }
        throw new Error('Could not read the selected file.');
      });
  }

  // ---------------------------------------------------------------------------
  // mapping.js
  // ---------------------------------------------------------------------------

  function blankMapping() {
    return { type: 'blank' };
  }

  function sourceMappingFromColumn(sourceColumn) {
    return {
      type: 'source',
      sourceKey: sourceColumn.key,
      sourceSection: sourceColumn.section,
      sourceField: sourceColumn.field,
      occurrence: sourceColumn.occurrence,
    };
  }

  function createEmptyMappingConfig() {
    return {
      configVersion: CONFIG_VERSION,
      applicationVersion: APP_VERSION,
      propertyMatch: blankMapping(),
      unitMatch: blankMapping(),
      columnMappings: STANDARD_TEMPLATE_COLUMNS.map(function (col) {
        return {
          targetColumnId: col.id,
          propertyMapping: blankMapping(),
          unitMapping: blankMapping(),
        };
      }),
    };
  }

  function cloneMappingConfig(config) {
    return JSON.parse(JSON.stringify(config));
  }

  function readEmbeddedConfig() {
    var el = document.getElementById('embedded-config');
    if (!el) return null;

    var raw = (el.textContent || '').trim();
    if (!raw) return null;

    try {
      return JSON.parse(raw);
    } catch (err) {
      console.error('Failed to parse embedded-config JSON:', err);
      return null;
    }
  }

  function normaliseSideMapping(mapping) {
    if (!mapping || typeof mapping !== 'object') return blankMapping();
    if (mapping.type === 'source' && typeof mapping.sourceKey === 'string' && mapping.sourceKey) {
      return {
        type: 'source',
        sourceKey: mapping.sourceKey,
        sourceSection: mapping.sourceSection != null ? mapping.sourceSection : '',
        sourceField: mapping.sourceField != null ? mapping.sourceField : '',
        occurrence: Number(mapping.occurrence) || 1,
      };
    }
    return blankMapping();
  }

  function normaliseMappingConfig(raw) {
    var base = createEmptyMappingConfig();
    if (!raw || typeof raw !== 'object') return base;

    if (raw.propertyMatch) {
      base.propertyMatch = normaliseSideMapping(raw.propertyMatch);
    }
    if (raw.unitMatch) {
      base.unitMatch = normaliseSideMapping(raw.unitMatch);
    }

    var byId = new Map();
    if (Array.isArray(raw.columnMappings)) {
      for (var i = 0; i < raw.columnMappings.length; i += 1) {
        var entry = raw.columnMappings[i];
        if (entry && typeof entry.targetColumnId === 'string') {
          byId.set(entry.targetColumnId, entry);
        }
      }
    }

    base.columnMappings = STANDARD_TEMPLATE_COLUMNS.map(function (col) {
      var existing = byId.get(col.id);
      if (!existing) {
        return {
          targetColumnId: col.id,
          propertyMapping: blankMapping(),
          unitMapping: blankMapping(),
        };
      }
      return {
        targetColumnId: col.id,
        propertyMapping: normaliseSideMapping(existing.propertyMapping),
        unitMapping: normaliseSideMapping(existing.unitMapping),
      };
    });

    base.configVersion = Number(raw.configVersion) || CONFIG_VERSION;
    base.applicationVersion = String(raw.applicationVersion || APP_VERSION);

    return base;
  }

  function resolveSourceKey(columns, sourceKey) {
    if (!columns || !sourceKey) return null;
    for (var i = 0; i < columns.length; i += 1) {
      if (columns[i].key === sourceKey) return columns[i];
    }
    return null;
  }

  function findExactSourceColumn(columns, section, field) {
    if (!columns) return null;
    var sectionKey = normaliseHeader(section);
    var fieldKey = normaliseHeader(field);
    for (var i = 0; i < columns.length; i += 1) {
      var c = columns[i];
      if (
        normaliseHeader(c.section) === sectionKey &&
        normaliseHeader(c.field) === fieldKey &&
        c.occurrence === 1
      ) {
        return c;
      }
    }
    return null;
  }

  function applyDefaultMatchFields(config, propertyParsed, unitParsed) {
    if (config.propertyMatch && config.propertyMatch.type !== 'source') {
      var propCol = findExactSourceColumn(
        propertyParsed && propertyParsed.columns,
        DEFAULT_MATCH_HINTS.property.section,
        DEFAULT_MATCH_HINTS.property.field
      );
      if (propCol) config.propertyMatch = sourceMappingFromColumn(propCol);
    }

    if (config.unitMatch && config.unitMatch.type !== 'source') {
      var unitCol = findExactSourceColumn(
        unitParsed && unitParsed.columns,
        DEFAULT_MATCH_HINTS.unit.section,
        DEFAULT_MATCH_HINTS.unit.field
      );
      if (unitCol) config.unitMatch = sourceMappingFromColumn(unitCol);
    }

    return config;
  }

  function resolveSideStatus(mapping, columns) {
    if (!mapping || mapping.type !== 'source') return 'blank';
    if (!columns) {
      return 'mapped';
    }
    return resolveSourceKey(columns, mapping.sourceKey) ? 'mapped' : 'missing';
  }

  function getColumnMappingStatus(entry, propertyParsed, unitParsed) {
    var propStatus = resolveSideStatus(entry.propertyMapping, propertyParsed && propertyParsed.columns);
    var unitStatus = resolveSideStatus(entry.unitMapping, unitParsed && unitParsed.columns);

    if (propStatus === 'missing' || unitStatus === 'missing') {
      return 'Missing source';
    }

    var propMapped = propStatus === 'mapped';
    var unitMapped = unitStatus === 'mapped';

    if (propMapped && unitMapped) return 'Mapped';
    if (propMapped || unitMapped) return 'Partially mapped';
    return 'Blank';
  }

  function getMappingSummary(config, propertyParsed, unitParsed) {
    var propertyMappings = 0;
    var unitMappings = 0;
    var missingMappings = 0;
    var mappedColumns = 0;
    var partiallyMapped = 0;
    var blankColumns = 0;

    for (var i = 0; i < config.columnMappings.length; i += 1) {
      var entry = config.columnMappings[i];
      var status = getColumnMappingStatus(entry, propertyParsed, unitParsed);
      if (entry.propertyMapping && entry.propertyMapping.type === 'source') propertyMappings += 1;
      if (entry.unitMapping && entry.unitMapping.type === 'source') unitMappings += 1;

      if (status === 'Missing source') missingMappings += 1;
      else if (status === 'Mapped') mappedColumns += 1;
      else if (status === 'Partially mapped') partiallyMapped += 1;
      else blankColumns += 1;
    }

    if (resolveSideStatus(config.propertyMatch, propertyParsed && propertyParsed.columns) === 'missing') {
      missingMappings += 1;
    }
    if (resolveSideStatus(config.unitMatch, unitParsed && unitParsed.columns) === 'missing') {
      missingMappings += 1;
    }

    return {
      standardColumns: STANDARD_TEMPLATE_COLUMNS.length,
      propertyMappings: propertyMappings,
      unitMappings: unitMappings,
      missingMappings: missingMappings,
      mappedColumns: mappedColumns,
      partiallyMapped: partiallyMapped,
      blankColumns: blankColumns,
    };
  }

  function isMappingConfigured(config, propertyParsed, unitParsed) {
    config = config != null ? config : APP_STATE.mappingConfig;
    propertyParsed = propertyParsed != null ? propertyParsed : APP_STATE.property;
    unitParsed = unitParsed != null ? unitParsed : APP_STATE.unit;

    if (!config) return false;

    if (config.propertyMatch && config.propertyMatch.type !== 'source') return false;
    if (config.unitMatch && config.unitMatch.type !== 'source') return false;

    if (!propertyParsed || !unitParsed) return false;

    var summary = getMappingSummary(config, propertyParsed, unitParsed);
    return summary.missingMappings === 0;
  }

  function getMainMappingStatusText() {
    var config = APP_STATE.mappingConfig;
    if (!config) return 'Not configured';

    var hasFiles = Boolean(APP_STATE.property && APP_STATE.unit);

    if (!hasFiles) {
      if (
        (config.propertyMatch && config.propertyMatch.type === 'source') ||
        (config.unitMatch && config.unitMatch.type === 'source')
      ) {
        return 'Configuration loaded';
      }
      var hasAny = config.columnMappings.some(function (e) {
        return (
          (e.propertyMapping && e.propertyMapping.type === 'source') ||
          (e.unitMapping && e.unitMapping.type === 'source')
        );
      });
      return hasAny ? 'Configuration loaded' : 'Not configured';
    }

    var summary = getMappingSummary(config, APP_STATE.property, APP_STATE.unit);
    if (summary.missingMappings > 0) {
      return 'Missing sources (' + summary.missingMappings + ')';
    }

    if (
      (config.propertyMatch && config.propertyMatch.type !== 'source') ||
      (config.unitMatch && config.unitMatch.type !== 'source')
    ) {
      return 'Match fields required';
    }

    return '\u2713 Configured';
  }

  function initialiseMappingState() {
    var embedded = readEmbeddedConfig();
    APP_STATE.mappingConfig = normaliseMappingConfig(embedded);
    return APP_STATE.mappingConfig;
  }

  function applyMappingConfig(config) {
    APP_STATE.mappingConfig = normaliseMappingConfig(config);
    APP_STATE.mappingConfigured = isMappingConfigured();
    return APP_STATE.mappingConfig;
  }

  function toEmbeddableConfig(config) {
    var normalised = normaliseMappingConfig(config);
    return {
      configVersion: CONFIG_VERSION,
      applicationVersion: APP_VERSION,
      propertyMatch: normalised.propertyMatch,
      unitMatch: normalised.unitMatch,
      columnMappings: normalised.columnMappings,
    };
  }

  function normaliseJoinValue(value) {
    return normaliseMatchingKey(value);
  }

  function getTemplateColumn(id) {
    for (var i = 0; i < STANDARD_TEMPLATE_COLUMNS.length; i += 1) {
      if (STANDARD_TEMPLATE_COLUMNS[i].id === id) return STANDARD_TEMPLATE_COLUMNS[i];
    }
    return null;
  }

  function getMappingConfig() {
    return APP_STATE.mappingConfig;
  }

  // ---------------------------------------------------------------------------
  // transformer.js
  // ---------------------------------------------------------------------------

  function resolveMappedValue(mapping, parsedRow, sourceColumns) {
    if (!mapping || mapping.type !== 'source' || !mapping.sourceKey) {
      return null;
    }

    var col = resolveSourceKey(sourceColumns, mapping.sourceKey);
    if (!col) return null;

    var raw = parsedRow && parsedRow.values && parsedRow.values[col.key];
    if (raw === undefined || raw === null || raw === '') return null;
    return raw;
  }

  function mapRowValues(kind, mappingConfig, parsedRow, parsedWorkbook) {
    var values = {};
    var side = kind === 'property' ? 'propertyMapping' : 'unitMapping';
    var columns = parsedWorkbook.columns;

    for (var i = 0; i < STANDARD_TEMPLATE_COLUMNS.length; i += 1) {
      var templateCol = STANDARD_TEMPLATE_COLUMNS[i];
      var entry = null;
      for (var j = 0; j < mappingConfig.columnMappings.length; j += 1) {
        if (mappingConfig.columnMappings[j].targetColumnId === templateCol.id) {
          entry = mappingConfig.columnMappings[j];
          break;
        }
      }
      var mapping = (entry && entry[side]) || { type: 'blank' };
      values[templateCol.id] = resolveMappedValue(mapping, parsedRow, columns);
    }

    return values;
  }

  function findDiagnosticFieldValue(parsedRow, sourceColumns, fieldNames) {
    var wanted = fieldNames.map(function (f) {
      return normaliseHeader(f);
    });
    var matches = sourceColumns.filter(function (c) {
      return wanted.indexOf(normaliseHeader(c.field)) !== -1;
    });
    if (matches.length === 0) return null;

    matches.sort(function (a, b) {
      var aGeneral = normaliseHeader(a.section) === 'general' ? 0 : 1;
      var bGeneral = normaliseHeader(b.section) === 'general' ? 0 : 1;
      if (aGeneral !== bGeneral) return aGeneral - bGeneral;
      if (a.occurrence !== b.occurrence) return a.occurrence - b.occurrence;
      return a.columnIndex - b.columnIndex;
    });

    var key = matches[0].key;
    var raw = parsedRow && parsedRow.values && parsedRow.values[key];
    if (raw === undefined || raw === null || raw === '') return null;
    return raw;
  }

  function rowHasMeaningfulData(row) {
    if (!row || !row.values) return false;
    var keys = Object.keys(row.values);
    for (var i = 0; i < keys.length; i += 1) {
      var value = row.values[keys[i]];
      if (value === null || value === undefined) continue;
      if (typeof value === 'string' && value.trim() === '') continue;
      return true;
    }
    return false;
  }

  function getActiveOutputColumns(mappingConfig) {
    var columns = [];
    if (!mappingConfig || !mappingConfig.columnMappings) return columns;

    for (var i = 0; i < STANDARD_TEMPLATE_COLUMNS.length; i += 1) {
      var templateCol = STANDARD_TEMPLATE_COLUMNS[i];
      var entry = null;
      for (var j = 0; j < mappingConfig.columnMappings.length; j += 1) {
        if (mappingConfig.columnMappings[j].targetColumnId === templateCol.id) {
          entry = mappingConfig.columnMappings[j];
          break;
        }
      }

      var propertyActive =
        entry && entry.propertyMapping && entry.propertyMapping.type === 'source';
      var unitActive = entry && entry.unitMapping && entry.unitMapping.type === 'source';

      // Both sides intentionally blank > omit column from generated workbook only.
      if (!propertyActive && !unitActive) continue;

      columns.push({ id: templateCol.id, label: templateCol.label });
    }

    return columns;
  }

  function validateForTransform(propertyData, unitData, mappingConfig) {
    var errors = [];
    var warnings = [];

    if (!propertyData) errors.push('Property Level workbook is not loaded.');
    if (!unitData) errors.push('Unit / TS Level workbook is not loaded.');
    if (!mappingConfig) errors.push('Mapping configuration is not available.');

    if (errors.length) {
      return { errors: errors, warnings: warnings, propertyMatchCol: null, unitMatchCol: null };
    }

    // Incomplete / unresolved mappings are warnings only - never block generation.
    var summary = getMappingSummary(mappingConfig, propertyData, unitData);
    if (summary.missingMappings > 0) {
      warnings.push(
        summary.missingMappings +
          ' configured source mapping(s) could not be resolved against the uploaded files. Those fields will be left blank.'
      );
    }

    var propertyMatchCol = null;
    var unitMatchCol = null;

    if (!mappingConfig.propertyMatch || mappingConfig.propertyMatch.type !== 'source') {
      warnings.push(
        'Property Level Match Field is not configured. Property summary rows will still be written, but Unit rows cannot be matched.'
      );
    } else {
      propertyMatchCol = resolveSourceKey(
        propertyData.columns,
        mappingConfig.propertyMatch.sourceKey
      );
      if (!propertyMatchCol) {
        warnings.push(
          'Property Level Match Field source could not be resolved. Property summary rows will still be written, but Unit rows cannot be matched.'
        );
      }
    }

    if (!mappingConfig.unitMatch || mappingConfig.unitMatch.type !== 'source') {
      warnings.push(
        'Unit / TS Match Field is not configured. Unit rows cannot be matched underneath properties.'
      );
    } else {
      unitMatchCol = resolveSourceKey(unitData.columns, mappingConfig.unitMatch.sourceKey);
      if (!unitMatchCol) {
        warnings.push(
          'Unit / TS Match Field source could not be resolved. Unit rows cannot be matched underneath properties.'
        );
      }
    }

    return {
      errors: errors,
      warnings: warnings,
      propertyMatchCol: propertyMatchCol,
      unitMatchCol: unitMatchCol,
    };
  }

  function transformToSchedule(propertyData, unitData, mappingConfig) {
    var validation = validateForTransform(propertyData, unitData, mappingConfig);
    var errors = validation.errors;
    var warnings = validation.warnings.slice();
    var propertyMatchCol = validation.propertyMatchCol;
    var unitMatchCol = validation.unitMatchCol;

    var columns = getActiveOutputColumns(mappingConfig);

    var emptySummary = {
      propertiesProcessed: 0,
      propertiesMatched: 0,
      propertiesWithoutMatchedUnits: 0,
      propertiesEmptyMatchValue: 0,
      unitRecordsProcessed: unitData && unitData.rows ? unitData.rows.length : 0,
      matchedUnitRecords: 0,
      unmatchedUnitRecords: 0,
      unitsEmptyMatchValue: 0,
      duplicatePropertyKeys: 0,
    };

    if (errors.length) {
      return {
        ok: false,
        errors: errors,
        warnings: warnings,
        summary: emptySummary,
        columns: columns,
        scheduleRows: [],
        unmatchedRows: [],
      };
    }

    // Valid property data rows only (ignore trailing blanks / empty non-data rows).
    var propertyRows = [];
    for (var pri = 0; pri < propertyData.rows.length; pri += 1) {
      if (rowHasMeaningfulData(propertyData.rows[pri])) {
        propertyRows.push(propertyData.rows[pri]);
      }
    }

    var keyIndex = new Map();
    var propertiesEmptyMatchValue = 0;

    if (propertyMatchCol) {
      for (var pki = 0; pki < propertyRows.length; pki += 1) {
        var prow = propertyRows[pki];
        var rawKey = prow.values[propertyMatchCol.key];
        var normKey = normaliseJoinValue(rawKey);
        if (!normKey) {
          propertiesEmptyMatchValue += 1;
          continue;
        }
        var keyEntry = keyIndex.get(normKey);
        if (!keyEntry) {
          keyEntry = { count: 0, displayValues: new Set(), rowNumbers: [] };
          keyIndex.set(normKey, keyEntry);
        }
        keyEntry.count += 1;
        keyEntry.rowNumbers.push(prow.rowIndex + 1);
        if (rawKey !== null && rawKey !== undefined && String(rawKey).trim() !== '') {
          keyEntry.displayValues.add(String(rawKey).trim());
        }
      }
    }

    var duplicateKeys = new Set();
    var duplicateDetails = [];
    keyIndex.forEach(function (v, k) {
      if (v.count > 1) {
        duplicateKeys.add(k);
        var label = v.displayValues.values().next().value || k;
        duplicateDetails.push('"' + label + '" (rows ' + v.rowNumbers.join(', ') + ')');
      }
    });

    if (propertiesEmptyMatchValue > 0) {
      warnings.push(
        propertiesEmptyMatchValue +
          ' Property row' +
          (propertiesEmptyMatchValue === 1 ? ' has' : 's have') +
          ' no match value. Their Property Summary rows will still be written with no Unit rows attached.'
      );
    }

    if (duplicateDetails.length > 0) {
      warnings.push(
        duplicateDetails.length +
          ' duplicate Property match key(s) found. Matching Unit records for those keys are treated as unmatched/ambiguous: ' +
          duplicateDetails.slice(0, 15).join('; ') +
          (duplicateDetails.length > 15
            ? ' \u2026and ' + (duplicateDetails.length - 15) + ' more.'
            : '')
      );
    }

    var unitsByKey = new Map();
    var unmatchedSourceRows = [];
    var unitsEmptyMatchValue = 0;
    var ambiguousUnitCount = 0;

    for (var ui = 0; ui < unitData.rows.length; ui += 1) {
      var unitRow = unitData.rows[ui];

      if (!unitMatchCol || !propertyMatchCol) {
        // Match fields unavailable/unresolved - cannot attach any units.
        unmatchedSourceRows.push(unitRow);
        continue;
      }

      var rawMatch = unitRow.values[unitMatchCol.key];
      var unitKey = normaliseJoinValue(rawMatch);

      if (!unitKey) {
        unitsEmptyMatchValue += 1;
        unmatchedSourceRows.push(unitRow);
        continue;
      }

      // Ambiguous when multiple properties share this key - do not attach to any.
      if (duplicateKeys.has(unitKey)) {
        ambiguousUnitCount += 1;
        unmatchedSourceRows.push(unitRow);
        continue;
      }

      var list = unitsByKey.get(unitKey);
      if (!list) {
        list = [];
        unitsByKey.set(unitKey, list);
      }
      list.push(unitRow);
    }

    if (unitsEmptyMatchValue > 0) {
      warnings.push(
        unitsEmptyMatchValue +
          ' Unit row' +
          (unitsEmptyMatchValue === 1 ? ' has' : 's have') +
          ' no match value.'
      );
    }

    var scheduleRows = [];
    var matchedUnitSet = new Set();
    var propertiesMatched = 0;

    for (var pi = 0; pi < propertyRows.length; pi += 1) {
      var propertyRow = propertyRows[pi];
      var propKey = '';
      if (propertyMatchCol) {
        propKey = normaliseJoinValue(propertyRow.values[propertyMatchCol.key]) || '';
      }

      scheduleRows.push({
        kind: 'property',
        values: mapRowValues('property', mappingConfig, propertyRow, propertyData),
        sourceRowIndex: propertyRow.rowIndex,
        matchKey: propKey,
      });

      var canAttach =
        Boolean(propKey) && !duplicateKeys.has(propKey) && unitsByKey.has(propKey);
      var units = canAttach ? unitsByKey.get(propKey) : [];

      if (units.length > 0) propertiesMatched += 1;

      for (var uj = 0; uj < units.length; uj += 1) {
        var matchedUnitRow = units[uj];
        matchedUnitSet.add(matchedUnitRow);
        scheduleRows.push({
          kind: 'unit',
          values: mapRowValues('unit', mappingConfig, matchedUnitRow, unitData),
          sourceRowIndex: matchedUnitRow.rowIndex,
          matchKey: propKey,
        });
      }
    }

    for (var uk = 0; uk < unitData.rows.length; uk += 1) {
      var checkRow = unitData.rows[uk];
      if (!matchedUnitSet.has(checkRow) && unmatchedSourceRows.indexOf(checkRow) === -1) {
        unmatchedSourceRows.push(checkRow);
      }
    }

    unmatchedSourceRows.sort(function (a, b) {
      return a.rowIndex - b.rowIndex;
    });

    var unmatchedRows = unmatchedSourceRows.map(function (row) {
      var rowMatch = unitMatchCol ? row.values[unitMatchCol.key] : null;
      return {
        originalExcelRow: row.rowIndex + 1,
        matchValue: rowMatch === undefined || rowMatch === '' ? null : rowMatch,
        propertyName: findDiagnosticFieldValue(row, unitData.columns, ['Property Name']),
        tenantName: findDiagnosticFieldValue(row, unitData.columns, ['Tenant Name']),
        suite: findDiagnosticFieldValue(row, unitData.columns, ['Suite']),
      };
    });

    if (unmatchedRows.length > 0) {
      warnings.push(
        unmatchedRows.length +
          ' Unit / TS record(s) did not match any Property' +
          (ambiguousUnitCount > 0
            ? ' (including ' + ambiguousUnitCount + ' ambiguous due to duplicate property keys)'
            : '') +
          ' and will be included on the "Unmatched Records" sheet.'
      );
    }

    var propertiesWithoutMatchedUnits = propertyRows.length - propertiesMatched;

    return {
      ok: true,
      errors: [],
      warnings: warnings,
      summary: {
        propertiesProcessed: propertyRows.length,
        propertiesMatched: propertiesMatched,
        propertiesWithoutMatchedUnits: propertiesWithoutMatchedUnits,
        propertiesEmptyMatchValue: propertiesEmptyMatchValue,
        unitRecordsProcessed: unitData.rows.length,
        matchedUnitRecords: matchedUnitSet.size,
        unmatchedUnitRecords: unmatchedRows.length,
        unitsEmptyMatchValue: unitsEmptyMatchValue,
        duplicatePropertyKeys: duplicateDetails.length,
      },
      columns: columns,
      scheduleRows: scheduleRows,
      unmatchedRows: unmatchedRows,
    };
  }

  /**
   * Lightweight matching analysis for the pre-generation Matching Summary UI.
   * Does not mutate configuration. Uses the same rules as transformToSchedule.
   */
  function analyseMatching(propertyData, unitData, mappingConfig) {
    if (!propertyData || !unitData || !mappingConfig) return null;
    var model = transformToSchedule(propertyData, unitData, mappingConfig);
    return {
      ok: model.ok,
      errors: model.errors,
      warnings: model.warnings,
      summary: model.summary,
    };
  }

  // ---------------------------------------------------------------------------
  // excelWriter.js
  // ---------------------------------------------------------------------------

  var HEADER_FILL = { patternType: 'solid', fgColor: { rgb: '002060' } };
  var HEADER_FONT = { bold: true, color: { rgb: 'FFFFFF' }, name: 'Calibri', sz: 11 };
  var PROPERTY_FILL = { patternType: 'solid', fgColor: { rgb: 'D9D9D9' } };
  var PROPERTY_FONT = { bold: true, name: 'Calibri', sz: 11, color: { rgb: '000000' } };
  var UNIT_FONT = { bold: false, name: 'Calibri', sz: 11, color: { rgb: '000000' } };

  var BORDER_COLOR = { rgb: 'BFBFBF' };
  var HEADER_BORDER_COLOR = { rgb: '001538' };

  var THIN_BORDER = {
    top: { style: 'thin', color: BORDER_COLOR },
    bottom: { style: 'thin', color: BORDER_COLOR },
    left: { style: 'thin', color: BORDER_COLOR },
    right: { style: 'thin', color: BORDER_COLOR },
  };

  var HEADER_BORDER = {
    top: { style: 'thin', color: HEADER_BORDER_COLOR },
    bottom: { style: 'thin', color: HEADER_BORDER_COLOR },
    left: { style: 'thin', color: HEADER_BORDER_COLOR },
    right: { style: 'thin', color: HEADER_BORDER_COLOR },
  };

  var PROPERTY_BORDER = {
    top: { style: 'medium', color: { rgb: '595959' } },
    bottom: { style: 'thin', color: BORDER_COLOR },
    left: { style: 'thin', color: BORDER_COLOR },
    right: { style: 'thin', color: BORDER_COLOR },
  };

  var COLUMN_NUM_FMT = {
    area_sqft: '#,##0',
    contracted_rent_pa: '\u00a3#,##0',
    contracted_rent_psf: '\u00a30.00',
    deemed_rent_psf: '\u00a30.00',
    market_rent_pa: '\u00a3#,##0',
    market_rent_psf: '\u00a30.00',
    rent_free_month: '0',
    expiry_void_month: '0',
    rent_free_expiry_month: '0',
    cap_rate: '0.00%',
    niy: '0.00%',
    ney: '0.00%',
    ry: '0.00%',
    cv_psf: '\u00a30',
    cap_ex: '\u00a3#,##0;[Red]-\u00a3#,##0',
    fair_value_april_2024: '\u00a3#,##0',
    fair_value_march_2026: '\u00a3#,##0',
    pct_change_26_april_24: '0.00%',
    pct_change_x_ytd: '0.00%',
    start_date: 'dd/mm/yyyy',
    review_date: 'dd/mm/yyyy',
    expiry_date: 'dd/mm/yyyy',
    break_date: 'dd/mm/yyyy',
  };

  var COLUMN_WIDTHS = {
    property_reference: 14,
    town: 14,
    address: 38,
    tenure_uk: 8,
    area_sqft: 12,
    unit: 30,
    tenant_name: 30,
    start_date: 12,
    review_date: 12,
    expiry_date: 12,
    break_date: 12,
    mwault_year: 10,
    contracted_rent_pa: 14,
    contracted_rent_psf: 12,
    deemed_rent_psf: 12,
    market_rent_pa: 14,
    market_rent_psf: 12,
    current_void: 11,
    rent_free_month: 10,
    expiry_void_month: 10,
    rent_free_expiry_month: 11,
    cap_rate: 9,
    niy: 8,
    ney: 8,
    ry: 8,
    cv_psf: 10,
    cap_ex: 12,
    fair_value_april_2024: 14,
    fair_value_march_2026: 14,
    pct_change_26_april_24: 12,
    pct_change_x_ytd: 11,
    comments: 36,
  };

  var DATE_COLUMN_IDS = new Set(['start_date', 'review_date', 'expiry_date', 'break_date']);
  var WRAP_COLUMN_IDS = new Set(['comments', 'address', 'unit', 'tenant_name']);

  function getNumberFormat(columnId, value) {
    if (columnId === 'current_void') {
      if (value instanceof Date && !Number.isNaN(value.getTime())) return 'dd/mm/yyyy';
      if (typeof value === 'number' && Number.isFinite(value)) return '0';
      return null;
    }

    if (COLUMN_NUM_FMT[columnId]) return COLUMN_NUM_FMT[columnId];

    if (value instanceof Date && !Number.isNaN(value.getTime())) return 'dd/mm/yyyy';

    return null;
  }

  function valueToCell(value, baseStyle, columnId) {
    baseStyle = baseStyle || {};
    columnId = columnId || '';
    var numFmt = getNumberFormat(columnId, value);
    var style = numFmt ? Object.assign({}, baseStyle, { numFmt: numFmt }) : Object.assign({}, baseStyle);

    if (value === null || value === undefined || value === '') {
      return { t: 's', v: '', s: style };
    }

    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      return {
        t: 'd',
        v: value,
        z: numFmt || 'dd/mm/yyyy',
        s: Object.assign({}, style, { numFmt: numFmt || 'dd/mm/yyyy' }),
      };
    }

    if (typeof value === 'number' && Number.isFinite(value)) {
      var cell = {
        t: 'n',
        v: value,
        s: style,
      };
      if (numFmt) cell.z = numFmt;
      return cell;
    }

    if (typeof value === 'boolean') {
      return { t: 'b', v: value, s: style };
    }

    return {
      t: 's',
      v: String(value),
      s: style,
    };
  }

  function buildDataCellStyle(columnId, isProperty) {
    var wrap = WRAP_COLUMN_IDS.has(columnId);
    var horizontal =
      DATE_COLUMN_IDS.has(columnId) || COLUMN_NUM_FMT[columnId] || columnId === 'current_void'
        ? 'right'
        : 'left';

    return {
      font: isProperty ? PROPERTY_FONT : UNIT_FONT,
      fill: isProperty ? PROPERTY_FILL : undefined,
      border: isProperty ? PROPERTY_BORDER : THIN_BORDER,
      alignment: {
        vertical: 'center',
        horizontal: horizontal,
        wrapText: wrap,
      },
    };
  }

  function buildTenancyScheduleSheet(model) {
    var columns = model.columns;
    var colCount = columns.length;
    var rowCount = 1 + model.scheduleRows.length;
    var ws = {};

    if (colCount === 0) {
      ws.A1 = {
        t: 's',
        v: 'No output columns configured (all Standardised Template mappings are blank).',
        s: {
          font: HEADER_FONT,
          fill: HEADER_FILL,
          border: HEADER_BORDER,
        },
      };
      ws['!ref'] = 'A1';
      ws['!cols'] = [{ wch: 80 }];
      return ws;
    }

    for (var c = 0; c < colCount; c += 1) {
      var headerAddr = XLSX.utils.encode_cell({ r: 0, c: c });
      ws[headerAddr] = {
        t: 's',
        v: columns[c].label,
        s: {
          fill: HEADER_FILL,
          font: HEADER_FONT,
          border: HEADER_BORDER,
          alignment: {
            vertical: 'center',
            horizontal: 'center',
            wrapText: true,
          },
        },
      };
    }

    for (var r = 0; r < model.scheduleRows.length; r += 1) {
      var row = model.scheduleRows[r];
      var isProperty = row.kind === 'property';

      for (var dc = 0; dc < colCount; dc += 1) {
        var colId = columns[dc].id;
        var dataAddr = XLSX.utils.encode_cell({ r: r + 1, c: dc });
        var style = buildDataCellStyle(colId, isProperty);
        if (!isProperty) delete style.fill;
        ws[dataAddr] = valueToCell(row.values[colId], style, colId);
      }
    }

    ws['!ref'] = XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: Math.max(rowCount - 1, 0), c: Math.max(colCount - 1, 0) },
    });

    ws['!freeze'] = {
      xSplit: 0,
      ySplit: 1,
      topLeftCell: 'A2',
      activePane: 'bottomLeft',
      state: 'frozen',
    };
    ws['!views'] = [{ state: 'frozen', ySplit: 1, topLeftCell: 'A2', activeCell: 'A2' }];

    if (colCount > 0) {
      ws['!autofilter'] = {
        ref: XLSX.utils.encode_range({
          s: { r: 0, c: 0 },
          e: { r: Math.max(rowCount - 1, 0), c: colCount - 1 },
        }),
      };
    }

    ws['!cols'] = columns.map(function (col) {
      return {
        wch:
          COLUMN_WIDTHS[col.id] != null
            ? COLUMN_WIDTHS[col.id]
            : Math.min(22, Math.max(10, String(col.label).length + 2)),
      };
    });

    var rowProps = [{ hpt: 36 }];
    for (var ri = 0; ri < model.scheduleRows.length; ri += 1) {
      rowProps.push({ hpt: 18 });
    }
    ws['!rows'] = rowProps;

    return ws;
  }

  function buildUnmatchedSheet(unmatchedRows) {
    var headers = [
      'Original Excel Row',
      'Configured Unit Match Value',
      'Property Name',
      'Tenant Name',
      'Suite',
    ];

    var ws = {};

    for (var c = 0; c < headers.length; c += 1) {
      var headerAddr = XLSX.utils.encode_cell({ r: 0, c: c });
      ws[headerAddr] = {
        t: 's',
        v: headers[c],
        s: {
          fill: HEADER_FILL,
          font: HEADER_FONT,
          border: HEADER_BORDER,
          alignment: { vertical: 'center', horizontal: 'center', wrapText: true },
        },
      };
    }

    for (var r = 0; r < unmatchedRows.length; r += 1) {
      var row = unmatchedRows[r];
      var values = [
        row.originalExcelRow,
        row.matchValue,
        row.propertyName,
        row.tenantName,
        row.suite,
      ];
      for (var dc = 0; dc < values.length; dc += 1) {
        var dataAddr = XLSX.utils.encode_cell({ r: r + 1, c: dc });
        var style = {
          font: UNIT_FONT,
          border: THIN_BORDER,
          alignment: { vertical: 'center', horizontal: dc === 0 ? 'right' : 'left' },
        };
        ws[dataAddr] = valueToCell(values[dc], style, dc === 0 ? 'area_sqft' : '');
      }
    }

    var lastRow = unmatchedRows.length;
    ws['!ref'] = XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: lastRow, c: headers.length - 1 },
    });
    ws['!views'] = [{ state: 'frozen', ySplit: 1, topLeftCell: 'A2', activeCell: 'A2' }];
    ws['!autofilter'] = {
      ref: XLSX.utils.encode_range({
        s: { r: 0, c: 0 },
        e: { r: lastRow, c: headers.length - 1 },
      }),
    };
    ws['!cols'] = [{ wch: 18 }, { wch: 32 }, { wch: 28 }, { wch: 28 }, { wch: 22 }];
    ws['!rows'] = [{ hpt: 28 }];

    return ws;
  }

  function buildWorkbookFromModel(model) {
    var wb = XLSX.utils.book_new();

    var scheduleSheet = buildTenancyScheduleSheet(model);
    XLSX.utils.book_append_sheet(wb, scheduleSheet, 'Tenancy Schedule');

    if (model.unmatchedRows && model.unmatchedRows.length > 0) {
      var unmatchedSheet = buildUnmatchedSheet(model.unmatchedRows);
      XLSX.utils.book_append_sheet(wb, unmatchedSheet, 'Unmatched Records');
    }

    return wb;
  }

  function workbookToArrayBuffer(workbook) {
    return XLSX.write(workbook, {
      bookType: 'xlsx',
      type: 'array',
      cellDates: true,
    });
  }

  function downloadWorkbook(workbook, filename) {
    var buffer = workbookToArrayBuffer(workbook);
    var blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.rel = 'noopener';
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(function () {
      URL.revokeObjectURL(url);
    }, 2000);
  }

  // ---------------------------------------------------------------------------
  // generator.js
  // ---------------------------------------------------------------------------

  var OUTPUT_FILENAME = 'Standardised_Tenancy_Schedule.xlsx';

  function generateStandardisedSchedule(propertyData, unitData, mappingConfig) {
    var model = transformToSchedule(propertyData, unitData, mappingConfig);

    if (!model.ok) {
      return {
        ok: false,
        errors: model.errors,
        warnings: model.warnings,
        summary: model.summary,
      };
    }

    try {
      var workbook = buildWorkbookFromModel(model);
      downloadWorkbook(workbook, OUTPUT_FILENAME);

      return {
        ok: true,
        errors: [],
        warnings: model.warnings,
        summary: model.summary,
        filename: OUTPUT_FILENAME,
      };
    } catch (err) {
      console.error('Workbook generation failed:', err);
      return {
        ok: false,
        errors: ['The Excel workbook could not be created. Please try again.'],
        warnings: model.warnings,
        summary: model.summary,
      };
    }
  }

  // ---------------------------------------------------------------------------
  // exportHtml.js
  // ---------------------------------------------------------------------------

  var EXPORT_FILENAME = 'Tenancy_Schedule_Standardiser_Configured.html';

  function serializeConfigForEmbed(data) {
    return JSON.stringify(data, null, 2)
      .replace(/</g, '\\u003c')
      .replace(/>/g, '\\u003e')
      .replace(/&/g, '\\u0026')
      .replace(/\u2028/g, '\\u2028')
      .replace(/\u2029/g, '\\u2029');
  }

  function isStandaloneDocument() {
    var externalScripts = Array.from(document.scripts).filter(function (s) {
      return s.src && !(s.textContent && s.textContent.trim());
    });
    return externalScripts.length === 0;
  }

  function ensureEmbeddedConfigElement(doc) {
    var el = doc.getElementById('embedded-config');
    if (el) return el;

    el = doc.createElement('script');
    el.id = 'embedded-config';
    el.type = 'application/json';
    var firstScript = doc.querySelector('script:not(#embedded-config)');
    if (firstScript && firstScript.parentNode) {
      firstScript.parentNode.insertBefore(el, firstScript);
    } else {
      (doc.body || doc.documentElement).appendChild(el);
    }
    return el;
  }

  function resetTransientUiState(doc) {
    var modal = doc.getElementById('settings-modal');
    if (modal) modal.classList.add('hidden');

    var confirmModal = doc.getElementById('confirm-modal');
    if (confirmModal) confirmModal.classList.add('hidden');

    doc.querySelectorAll('.dropzone').forEach(function (dz) {
      dz.classList.remove('has-file', 'has-error', 'drag-over');
      var idle = dz.querySelector('[data-role="idle"]');
      var success = dz.querySelector('[data-role="success"]');
      var error = dz.querySelector('[data-role="error"]');
      if (idle) idle.classList.remove('hidden');
      if (success) success.classList.add('hidden');
      if (error) {
        error.classList.add('hidden');
        error.textContent = '';
      }
      var filename = dz.querySelector('[data-role="filename"]');
      var columns = dz.querySelector('[data-role="columns"]');
      var rows = dz.querySelector('[data-role="rows"]');
      if (filename) filename.textContent = '';
      if (columns) columns.textContent = '';
      if (rows) rows.textContent = '';
    });

    doc.querySelectorAll('input[type="file"]').forEach(function (input) {
      input.removeAttribute('value');
      var clean = doc.createElement('input');
      clean.type = 'file';
      clean.accept = input.accept;
      clean.className = input.className;
      clean.hidden = input.hidden;
      if (input.dataset.role) clean.dataset.role = input.dataset.role;
      input.replaceWith(clean);
    });

    var propertyStatus = doc.getElementById('property-status');
    var unitStatus = doc.getElementById('unit-status');
    var mappingStatus = doc.getElementById('mapping-status');
    if (propertyStatus) propertyStatus.textContent = 'No file loaded';
    if (unitStatus) unitStatus.textContent = 'No file loaded';
    if (mappingStatus) mappingStatus.textContent = 'Configuration loaded';

    var generateBtn = doc.getElementById('generate-btn');
    if (generateBtn) generateBtn.setAttribute('disabled', '');

    var generateResult = doc.getElementById('generate-result');
    if (generateResult) {
      generateResult.classList.add('hidden');
      generateResult.classList.remove('result-success', 'result-error', 'result-warning');
      generateResult.innerHTML = '';
    }

    var matchingSummary = doc.getElementById('matching-summary');
    if (matchingSummary) {
      matchingSummary.classList.add('hidden');
      matchingSummary.classList.remove('summary-ok', 'summary-warn');
      matchingSummary.innerHTML = '';
    }

    var compareResults = doc.getElementById('compare-results');
    if (compareResults) {
      compareResults.classList.add('hidden');
    }
    var compareError = doc.getElementById('compare-error');
    if (compareError) {
      compareError.classList.add('hidden');
      compareError.textContent = '';
    }
    var compareSearch = doc.getElementById('compare-search');
    if (compareSearch) compareSearch.value = '';
    doc.querySelectorAll('#view-compare .dropzone').forEach(function (dz) {
      dz.classList.remove('has-file', 'has-error', 'drag-over');
      var idle = dz.querySelector('[data-role="idle"]');
      var success = dz.querySelector('[data-role="success"]');
      var error = dz.querySelector('[data-role="error"]');
      if (idle) idle.classList.remove('hidden');
      if (success) success.classList.add('hidden');
      if (error) {
        error.classList.add('hidden');
        error.textContent = '';
      }
    });
    var stdView = doc.getElementById('view-standardiser');
    var cmpView = doc.getElementById('view-compare');
    if (stdView) stdView.classList.remove('hidden');
    if (cmpView) cmpView.classList.add('hidden');
    var appEl = doc.getElementById('app');
    if (appEl) {
      appEl.classList.add('app-mode-standardiser');
      appEl.classList.remove('app-mode-compare');
    }

    var generateHint = doc.getElementById('generate-hint');
    if (generateHint) {
      generateHint.textContent =
        'Configure column mappings from Advanced Settings before generating the schedule.';
    }

    doc.querySelectorAll('[data-role="settings-feedback"]').forEach(function (el) {
      el.textContent = '';
      el.classList.add('hidden');
    });
  }

  function serializeDocument(doc) {
    var doctype = doc.doctype
      ? '<!DOCTYPE ' +
        doc.doctype.name +
        (doc.doctype.publicId ? ' PUBLIC "' + doc.doctype.publicId + '"' : '') +
        (!doc.doctype.publicId && doc.doctype.systemId ? ' SYSTEM' : '') +
        (doc.doctype.systemId ? ' "' + doc.doctype.systemId + '"' : '') +
        '>\n'
      : '<!DOCTYPE html>\n';

    return doctype + doc.documentElement.outerHTML;
  }

  function downloadHtml(html, filename) {
    var blob = new Blob([html], { type: 'text/html;charset=utf-8' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.rel = 'noopener';
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(function () {
      URL.revokeObjectURL(url);
    }, 2000);
  }

  function ensureUtf8CharsetMeta(doc) {
    var head = doc.head || doc.getElementsByTagName('head')[0];
    if (!head) return;
    var existing = head.querySelector('meta[charset]');
    if (existing) {
      existing.setAttribute('charset', 'UTF-8');
      if (head.firstChild !== existing) head.insertBefore(existing, head.firstChild);
      return;
    }
    var meta = doc.createElement('meta');
    meta.setAttribute('charset', 'UTF-8');
    head.insertBefore(meta, head.firstChild);
  }

  function exportConfiguredHtml(mappingConfig) {
    if (!isStandaloneDocument()) {
      return {
        ok: false,
        error:
          'Export Configured HTML requires a self-contained document without external script sources.',
      };
    }

    try {
      var payload = toEmbeddableConfig(mappingConfig);
      var jsonText = serializeConfigForEmbed(payload);

      var liveHtml = serializeDocument(document);
      var scratch = new DOMParser().parseFromString(liveHtml, 'text/html');

      var embedded = ensureEmbeddedConfigElement(scratch);
      embedded.textContent = jsonText;

      ensureUtf8CharsetMeta(scratch);
    resetTransientUiState(scratch);
      scratch.title = 'Tenancy Schedule Standardiser (Configured)';

      var html = serializeDocument(scratch);
      downloadHtml(html, EXPORT_FILENAME);

      return { ok: true, filename: EXPORT_FILENAME };
    } catch (err) {
      console.error('Export configured HTML failed:', err);
      return {
        ok: false,
        error: 'Could not create the configured HTML file. Please try again.',
      };
    }
  }

  // ---------------------------------------------------------------------------
  // settingsUI.js
  // ---------------------------------------------------------------------------

  var draftConfig = null;
  var onConfigChanged = null;

  function setupSettingsUI(options) {
    options = options || {};
    onConfigChanged = options.onConfigChanged || null;

    var modal = document.getElementById('settings-modal');
    var openBtn = document.getElementById('settings-btn');

    openBtn.addEventListener('click', function () {
      openSettings();
    });

    modal.querySelectorAll('[data-role="close"]').forEach(function (el) {
      el.addEventListener('click', function () {
        closeSettings();
      });
    });

    document.addEventListener('keydown', function (event) {
      if (event.key === 'Escape' && !modal.classList.contains('hidden')) {
        var confirmModal = document.getElementById('confirm-modal');
        if (confirmModal && !confirmModal.classList.contains('hidden')) {
          closeConfirm();
          return;
        }
        closeSettings();
      }
    });

    modal.querySelectorAll('[data-settings-tab]').forEach(function (tabBtn) {
      tabBtn.addEventListener('click', function () {
        activateTab(tabBtn.dataset.settingsTab);
      });
    });

    document.getElementById('btn-apply-config').addEventListener('click', function () {
      syncDraftFromDom();
      applyMappingConfig(draftConfig);
      showFeedback('Configuration applied to this session.', 'success');
      notifyChanged();
      renderAll();
    });

    document.getElementById('btn-export-html').addEventListener('click', function () {
      syncDraftFromDom();
      applyMappingConfig(draftConfig);
      notifyChanged();

      var result = exportConfiguredHtml(draftConfig);
      if (result.ok) {
        showFeedback(
          'Downloaded ' + result.filename + '. Open that file to use the embedded configuration.',
          'success'
        );
      } else {
        showFeedback(result.error, 'error');
      }
    });

    document.getElementById('btn-reset-mapping').addEventListener('click', function () {
      openConfirm(
        'Reset mapping?',
        'This clears all column mappings and match fields for the current session only. It does not modify any HTML file on disk.',
        function () {
          draftConfig = createEmptyMappingConfig();
          applyDefaultMatchFields(draftConfig, APP_STATE.property, APP_STATE.unit);
          applyMappingConfig(draftConfig);
          showFeedback('Mappings reset for this session.', 'success');
          notifyChanged();
          renderAll();
        }
      );
    });

    var filterInput = document.getElementById('mapping-filter');
    if (filterInput) {
      filterInput.addEventListener('input', function () {
        filterMappingRows(filterInput.value);
      });
    }
  }

  function notifyChanged() {
    if (typeof onConfigChanged === 'function') onConfigChanged();
  }

  function openSettings() {
    draftConfig = cloneMappingConfig(APP_STATE.mappingConfig || createEmptyMappingConfig());
    applyDefaultMatchFields(draftConfig, APP_STATE.property, APP_STATE.unit);

    var modal = document.getElementById('settings-modal');
    modal.classList.remove('hidden');
    activateTab('mapping');
    clearFeedback();
    renderAll();
  }

  function closeSettings() {
    document.getElementById('settings-modal').classList.add('hidden');
    draftConfig = null;
  }

  function activateTab(tabId) {
    var modal = document.getElementById('settings-modal');
    modal.querySelectorAll('[data-settings-tab]').forEach(function (btn) {
      btn.classList.toggle('active', btn.dataset.settingsTab === tabId);
    });
    modal.querySelectorAll('[data-settings-panel]').forEach(function (panel) {
      panel.classList.toggle('hidden', panel.dataset.settingsPanel !== tabId);
    });
  }

  function renderAll() {
    if (!draftConfig) return;
    renderMappingSummary();
    renderMappingTable();
    renderMatchFields();
    renderConfigTab();
  }

  function showFeedback(message, kind) {
    var el = document.querySelector('#settings-modal [data-role="settings-feedback"]');
    if (!el) return;
    el.textContent = message;
    el.classList.remove('hidden', 'feedback-success', 'feedback-error');
    el.classList.add(kind === 'error' ? 'feedback-error' : 'feedback-success');
  }

  function clearFeedback() {
    var el = document.querySelector('#settings-modal [data-role="settings-feedback"]');
    if (!el) return;
    el.textContent = '';
    el.classList.add('hidden');
    el.classList.remove('feedback-success', 'feedback-error');
  }

  function renderMappingSummary() {
    var summary = getMappingSummary(draftConfig, APP_STATE.property, APP_STATE.unit);
    var el = document.getElementById('mapping-summary');
    el.innerHTML =
      '<span><strong>' +
      summary.standardColumns +
      '</strong> Standard Columns</span>' +
      '<span><strong>' +
      summary.propertyMappings +
      '</strong> Property mappings</span>' +
      '<span><strong>' +
      summary.unitMappings +
      '</strong> Unit mappings</span>' +
      '<span class="' +
      (summary.missingMappings ? 'summary-missing' : '') +
      '">' +
      '<strong>' +
      summary.missingMappings +
      '</strong> Missing mappings</span>';
  }

  function buildSourceOptionsHtml(fileType, selectedMapping) {
    var parsed = APP_STATE[fileType];
    var columns = (parsed && parsed.columns) || [];
    var selectedKey =
      selectedMapping && selectedMapping.type === 'source' ? selectedMapping.sourceKey : '';

    var html = '<option value="">-- Blank --</option>';

    if (selectedKey && !columns.some(function (c) {
      return c.key === selectedKey;
    })) {
      var label = [
        selectedMapping.sourceSection || '(No Section)',
        selectedMapping.sourceField || '(unknown)',
      ].join(' > ');
      html +=
        '<option value="' +
        escapeAttr(selectedKey) +
        '" selected>\n      \u26a0 Missing: ' +
        escapeHtml(label) +
        '\n    </option>';
    }

    for (var i = 0; i < columns.length; i += 1) {
      var col = columns[i];
      var selected = col.key === selectedKey ? ' selected' : '';
      html +=
        '<option value="' +
        escapeAttr(col.key) +
        '"' +
        selected +
        '>' +
        escapeHtml(col.displayLabel) +
        '</option>';
    }

    if (!parsed) {
      html +=
        '<option value="" disabled>Upload a ' +
        (fileType === 'property' ? 'Property' : 'Unit') +
        ' file to see source fields</option>';
    }

    return html;
  }

  function renderMappingTable() {
    var tbody = document.getElementById('mapping-table-body');
    var filterValue = ((document.getElementById('mapping-filter') &&
      document.getElementById('mapping-filter').value) ||
      ''
    ).trim().toLowerCase();

    var rows = STANDARD_TEMPLATE_COLUMNS.map(function (col) {
      var entry = null;
      for (var i = 0; i < draftConfig.columnMappings.length; i += 1) {
        if (draftConfig.columnMappings[i].targetColumnId === col.id) {
          entry = draftConfig.columnMappings[i];
          break;
        }
      }
      if (!entry) {
        entry = {
          targetColumnId: col.id,
          propertyMapping: blankMapping(),
          unitMapping: blankMapping(),
        };
      }

      var status = getColumnMappingStatus(entry, APP_STATE.property, APP_STATE.unit);
      var statusClass = statusClassName(status);
      var hidden = filterValue && col.label.toLowerCase().indexOf(filterValue) === -1 ? 'hidden' : '';

      return (
        '<tr data-target-id="' +
        escapeAttr(col.id) +
        '" class="' +
        hidden +
        '">' +
        '<td class="col-standard">' +
        '<div class="standard-label">' +
        escapeHtml(col.label) +
        '</div>' +
        '<div class="standard-id">' +
        escapeHtml(col.id) +
        '</div>' +
        '</td>' +
        '<td>' +
        '<label class="sr-only" for="prop-map-' +
        escapeAttr(col.id) +
        '">Property Row Source for ' +
        escapeHtml(col.label) +
        '</label>' +
        '<select id="prop-map-' +
        escapeAttr(col.id) +
        '" data-map-side="property" data-target-id="' +
        escapeAttr(col.id) +
        '" class="source-select">' +
        buildSourceOptionsHtml('property', entry.propertyMapping) +
        '</select>' +
        '</td>' +
        '<td>' +
        '<label class="sr-only" for="unit-map-' +
        escapeAttr(col.id) +
        '">Unit Row Source for ' +
        escapeHtml(col.label) +
        '</label>' +
        '<select id="unit-map-' +
        escapeAttr(col.id) +
        '" data-map-side="unit" data-target-id="' +
        escapeAttr(col.id) +
        '" class="source-select">' +
        buildSourceOptionsHtml('unit', entry.unitMapping) +
        '</select>' +
        '</td>' +
        '<td class="col-status">' +
        '<span class="status-pill ' +
        statusClass +
        '">' +
        escapeHtml(status) +
        '</span>' +
        '</td>' +
        '</tr>'
      );
    }).join('');

    tbody.innerHTML = rows;

    tbody.querySelectorAll('select.source-select').forEach(function (select) {
      select.addEventListener('change', function () {
        var targetId = select.dataset.targetId;
        var side = select.dataset.mapSide;
        var key = select.value;
        updateDraftColumnMapping(targetId, side, key);
        renderMappingSummary();
        updateRowStatus(targetId);
      });
    });
  }

  function updateRowStatus(targetId) {
    var entry = null;
    for (var i = 0; i < draftConfig.columnMappings.length; i += 1) {
      if (draftConfig.columnMappings[i].targetColumnId === targetId) {
        entry = draftConfig.columnMappings[i];
        break;
      }
    }
    if (!entry) return;
    var status = getColumnMappingStatus(entry, APP_STATE.property, APP_STATE.unit);
    var row = document.querySelector(
      '#mapping-table-body tr[data-target-id="' + cssEscape(targetId) + '"]'
    );
    if (!row) return;
    var pill = row.querySelector('.status-pill');
    if (pill) {
      pill.textContent = status;
      pill.className = 'status-pill ' + statusClassName(status);
    }
  }

  function filterMappingRows(query) {
    var q = query.trim().toLowerCase();
    document.querySelectorAll('#mapping-table-body tr[data-target-id]').forEach(function (row) {
      var label = (row.querySelector('.standard-label') &&
        row.querySelector('.standard-label').textContent) ||
        '';
      row.classList.toggle('hidden', Boolean(q) && label.toLowerCase().indexOf(q) === -1);
    });
  }

  function updateDraftColumnMapping(targetId, side, sourceKey) {
    var entry = null;
    for (var i = 0; i < draftConfig.columnMappings.length; i += 1) {
      if (draftConfig.columnMappings[i].targetColumnId === targetId) {
        entry = draftConfig.columnMappings[i];
        break;
      }
    }
    if (!entry) {
      entry = {
        targetColumnId: targetId,
        propertyMapping: blankMapping(),
        unitMapping: blankMapping(),
      };
      draftConfig.columnMappings.push(entry);
    }

    var field = side === 'property' ? 'propertyMapping' : 'unitMapping';
    entry[field] = mappingFromSelectValue(side, sourceKey, entry[field]);
  }

  function mappingFromSelectValue(fileType, sourceKey, previous) {
    if (!sourceKey) return blankMapping();

    var cols = (APP_STATE[fileType] && APP_STATE[fileType].columns) || [];
    var col = null;
    for (var i = 0; i < cols.length; i += 1) {
      if (cols[i].key === sourceKey) {
        col = cols[i];
        break;
      }
    }
    if (col) return sourceMappingFromColumn(col);

    if (previous && previous.type === 'source' && previous.sourceKey === sourceKey) {
      return {
        type: 'source',
        sourceKey: sourceKey,
        sourceSection: previous.sourceSection || '',
        sourceField: previous.sourceField || '',
        occurrence: previous.occurrence || 1,
      };
    }

    return {
      type: 'source',
      sourceKey: sourceKey,
      sourceSection: '',
      sourceField: '',
      occurrence: 1,
    };
  }

  function renderMatchFields() {
    var propSelect = document.getElementById('property-match-select');
    var unitSelect = document.getElementById('unit-match-select');

    propSelect.innerHTML = buildSourceOptionsHtml('property', draftConfig.propertyMatch);
    unitSelect.innerHTML = buildSourceOptionsHtml('unit', draftConfig.unitMatch);

    propSelect.onchange = function () {
      draftConfig.propertyMatch = mappingFromSelectValue(
        'property',
        propSelect.value,
        draftConfig.propertyMatch
      );
    };
    unitSelect.onchange = function () {
      draftConfig.unitMatch = mappingFromSelectValue(
        'unit',
        unitSelect.value,
        draftConfig.unitMatch
      );
    };
  }

  function renderConfigTab() {
    document.getElementById('config-version-value').textContent = String(
      draftConfig.configVersion || CONFIG_VERSION
    );
    document.getElementById('app-version-value').textContent = APP_VERSION;
  }

  function syncDraftFromDom() {
    if (!draftConfig) return;

    document.querySelectorAll('#mapping-table-body select.source-select').forEach(function (select) {
      updateDraftColumnMapping(select.dataset.targetId, select.dataset.mapSide, select.value);
    });

    var propSelect = document.getElementById('property-match-select');
    var unitSelect = document.getElementById('unit-match-select');
    if (propSelect) {
      draftConfig.propertyMatch = mappingFromSelectValue(
        'property',
        propSelect.value,
        draftConfig.propertyMatch
      );
    }
    if (unitSelect) {
      draftConfig.unitMatch = mappingFromSelectValue(
        'unit',
        unitSelect.value,
        draftConfig.unitMatch
      );
    }
  }

  function refreshSettingsIfOpen() {
    var modal = document.getElementById('settings-modal');
    if (!modal || modal.classList.contains('hidden') || !draftConfig) return;

    syncDraftFromDom();
    applyDefaultMatchFields(draftConfig, APP_STATE.property, APP_STATE.unit);
    renderAll();
  }

  function statusClassName(status) {
    switch (status) {
      case 'Mapped':
        return 'status-mapped';
      case 'Partially mapped':
        return 'status-partial';
      case 'Missing source':
        return 'status-missing';
      default:
        return 'status-blank';
    }
  }

  function openConfirm(title, message, onConfirm) {
    var modal = document.getElementById('confirm-modal');
    document.getElementById('confirm-title').textContent = title;
    document.getElementById('confirm-message').textContent = message;
    modal.classList.remove('hidden');

    var okBtn = document.getElementById('confirm-ok');
    var cancelBtn = document.getElementById('confirm-cancel');

    var cleanup = function () {
      okBtn.onclick = null;
      cancelBtn.onclick = null;
      modal.querySelectorAll('[data-role="confirm-close"]').forEach(function (el) {
        el.onclick = null;
      });
    };

    var close = function () {
      cleanup();
      closeConfirm();
    };

    okBtn.onclick = function () {
      close();
      onConfirm();
    };
    cancelBtn.onclick = close;
    modal.querySelectorAll('[data-role="confirm-close"]').forEach(function (el) {
      el.onclick = close;
    });
  }

  function closeConfirm() {
    document.getElementById('confirm-modal').classList.add('hidden');
  }

  function escapeHtml(value) {
    return String(value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function escapeAttr(value) {
    return escapeHtml(value).replace(/'/g, '&#39;');
  }

  function cssEscape(value) {
    if (typeof CSS !== 'undefined' && CSS.escape) return CSS.escape(value);
    return String(value).replace(/["\\]/g, '\\$&');
  }

  // ---------------------------------------------------------------------------
  // main.js
  // ---------------------------------------------------------------------------

  function formatCount(n) {
    return Number(n).toLocaleString('en-GB');
  }

  function qs(root, selector) {
    return root.querySelector(selector);
  }

  function setStatusLines(el, lines) {
    if (!el) return;
    el.innerHTML =
      '<div class="status-lines">' +
      lines
        .map(function (line) {
          return '<div>' + escapeHtml(line) + '</div>';
        })
        .join('') +
      '</div>';
  }

  function getFileStatusLines(parsed, singular, plural) {
    if (!parsed) return ['No file loaded'];
    var n = parsed.dataRowCount;
    return [
      '\u2713 ' + parsed.fileName,
      formatCount(n) + ' ' + (n === 1 ? singular : plural) + ' detected',
      formatCount(parsed.columns.length) + ' source columns detected',
    ];
  }

  function getMappingStatusLines() {
    var config = APP_STATE.mappingConfig;
    if (!config) return ['Not configured'];

    var hasFiles = Boolean(APP_STATE.property && APP_STATE.unit);

    if (!hasFiles) {
      if (
        (config.propertyMatch && config.propertyMatch.type === 'source') ||
        (config.unitMatch && config.unitMatch.type === 'source')
      ) {
        return ['Configuration loaded'];
      }
      var hasAny = config.columnMappings.some(function (e) {
        return (
          (e.propertyMapping && e.propertyMapping.type === 'source') ||
          (e.unitMapping && e.unitMapping.type === 'source')
        );
      });
      return [hasAny ? 'Configuration loaded' : 'Not configured'];
    }

    if (APP_STATE.mappingConfigured) {
      return ['\u2713 Configured'];
    }

    return ['\u26a0 Mapping requires attention'];
  }

  function setupDropzone(dropzone) {
    var fileType = dropzone.dataset.fileType;
    var input = qs(dropzone, '[data-role="input"]');
    var browseBtn = qs(dropzone, '[data-role="browse"]');
    var idle = qs(dropzone, '[data-role="idle"]');
    var success = qs(dropzone, '[data-role="success"]');
    var errorEl = qs(dropzone, '[data-role="error"]');
    var filenameEl = qs(dropzone, '[data-role="filename"]');
    var columnsEl = qs(dropzone, '[data-role="columns"]');
    var rowsEl = qs(dropzone, '[data-role="rows"]');

    function showIdle() {
      idle.classList.remove('hidden');
      success.classList.add('hidden');
      errorEl.classList.add('hidden');
      dropzone.classList.remove('has-file', 'has-error');
    }

    function showError(message) {
      idle.classList.add('hidden');
      success.classList.add('hidden');
      errorEl.classList.remove('hidden');
      errorEl.textContent = message;
      dropzone.classList.remove('has-file');
      dropzone.classList.add('has-error');
    }

    function showSuccess(parsed) {
      idle.classList.add('hidden');
      errorEl.classList.add('hidden');
      success.classList.remove('hidden');
      filenameEl.textContent = '\u2713 ' + parsed.fileName;
      columnsEl.textContent =
        '\u2713 ' + formatCount(parsed.columns.length) + ' source columns detected';
      rowsEl.textContent =
        '\u2713 ' + formatCount(parsed.dataRowCount) + ' data rows detected';
      dropzone.classList.add('has-file');
      dropzone.classList.remove('has-error');
    }

    function handleFile(file) {
      if (!file) return;

      parseSourceFile(file, fileType)
        .then(function (parsed) {
          APP_STATE[fileType] = parsed;

          if (APP_STATE.mappingConfig) {
            applyDefaultMatchFields(APP_STATE.mappingConfig, APP_STATE.property, APP_STATE.unit);
          }

          showSuccess(parsed);
          clearGenerateResult();
          refreshStatus();
          refreshSettingsIfOpen();
        })
        .catch(function (err) {
          APP_STATE[fileType] = null;
          var message =
            err && typeof err.message === 'string' && err.message
              ? err.message
              : 'The file could not be processed. Please check it is a valid .xlsx export.';
          showError(message);
          clearGenerateResult();
          refreshStatus();
          refreshSettingsIfOpen();
          console.error('[' + fileType + '] parse failed:', err);
        });
    }

    function openPicker(event) {
      if (event) {
        event.stopPropagation();
      }
      // Do not preventDefault ? some browsers suppress the file dialog if the initiating click is cancelled.
      if (!input) return;
      try {
        input.click();
      } catch (err) {
        console.error('File picker could not be opened:', err);
      }
    }

    if (browseBtn) {
      browseBtn.addEventListener('click', function (event) {
        event.stopPropagation();
        openPicker(event);
      });
    }

    dropzone.addEventListener('click', function (event) {
      if (browseBtn && (event.target === browseBtn || browseBtn.contains(event.target))) return;
      if (event.target && event.target.tagName === 'INPUT') return;
      openPicker(event);
    });

    dropzone.addEventListener('keydown', function (event) {
      if (event.key === 'Enter' || event.key === ' ') {
        openPicker(event);
      }
    });

    input.addEventListener('change', function () {
      var file = input.files && input.files[0];
      input.value = '';
      handleFile(file);
    });

    dropzone.addEventListener('dragenter', function (event) {
      event.preventDefault();
      dropzone.classList.add('drag-over');
    });

    dropzone.addEventListener('dragover', function (event) {
      event.preventDefault();
      dropzone.classList.add('drag-over');
    });

    dropzone.addEventListener('dragleave', function (event) {
      if (!dropzone.contains(event.relatedTarget)) {
        dropzone.classList.remove('drag-over');
      }
    });

    dropzone.addEventListener('drop', function (event) {
      event.preventDefault();
      dropzone.classList.remove('drag-over');
      var file =
        event.dataTransfer && event.dataTransfer.files && event.dataTransfer.files[0];
      handleFile(file);
    });

    showIdle();
  }

  function clearGenerateResult() {
    var el = document.getElementById('generate-result');
    if (!el) return;
    el.classList.add('hidden');
    el.classList.remove('result-success', 'result-error', 'result-warning');
    el.innerHTML = '';
  }

  function renderMatchingSummaryPanel(analysis) {
    var el = document.getElementById('matching-summary');
    if (!el) return;

    if (!analysis || !analysis.summary) {
      el.classList.add('hidden');
      el.classList.remove('summary-ok', 'summary-warn');
      el.innerHTML = '';
      return;
    }

    var s = analysis.summary;
    var hasIssues =
      s.propertiesWithoutMatchedUnits > 0 ||
      s.unmatchedUnitRecords > 0 ||
      s.propertiesEmptyMatchValue > 0 ||
      s.unitsEmptyMatchValue > 0 ||
      s.duplicatePropertyKeys > 0;

    el.classList.remove('hidden', 'summary-ok', 'summary-warn');
    el.classList.add(hasIssues ? 'summary-warn' : 'summary-ok');

    var statusLine = hasIssues
      ? '\u26a0 Schedule will be generated with unmatched records'
      : '\u2713 Schedule can be generated';

    var secondary = [];
    if (s.propertiesEmptyMatchValue > 0) {
      secondary.push(
        s.propertiesEmptyMatchValue +
          ' Property row' +
          (s.propertiesEmptyMatchValue === 1 ? ' has' : 's have') +
          ' no match value'
      );
    }
    if (s.unitsEmptyMatchValue > 0) {
      secondary.push(
        s.unitsEmptyMatchValue +
          ' Unit row' +
          (s.unitsEmptyMatchValue === 1 ? ' has' : 's have') +
          ' no match value'
      );
    }
    if (s.duplicatePropertyKeys > 0) {
      secondary.push(
        s.duplicatePropertyKeys +
          ' duplicate Property match key' +
          (s.duplicatePropertyKeys === 1 ? '' : 's')
      );
    }

    el.innerHTML =
      '<h3>Matching Summary</h3>' +
      '<div class="summary-block">' +
      '<strong>Property Level</strong><br>' +
      formatCount(s.propertiesProcessed) +
      ' properties detected<br>' +
      formatCount(s.propertiesMatched) +
      ' properties matched to one or more TS records<br>' +
      formatCount(s.propertiesWithoutMatchedUnits) +
      ' properties with no matched TS records' +
      '</div>' +
      '<div class="summary-block">' +
      '<strong>Unit / TS Level</strong><br>' +
      formatCount(s.unitRecordsProcessed) +
      ' unit records detected<br>' +
      formatCount(s.matchedUnitRecords) +
      ' unit records matched<br>' +
      formatCount(s.unmatchedUnitRecords) +
      ' unit records unmatched' +
      '</div>' +
      '<p class="summary-status">' +
      statusLine +
      '</p>' +
      (secondary.length
        ? '<ul class="summary-secondary">' +
          secondary
            .map(function (line) {
              return '<li>' + escapeHtml(line) + '</li>';
            })
            .join('') +
          '</ul>'
        : '');
  }

  function showGenerateResult(result) {
    var el = document.getElementById('generate-result');
    if (!el) return;

    el.classList.remove('hidden', 'result-success', 'result-error', 'result-warning');

    var summary = result.summary || {
      propertiesProcessed: 0,
      propertiesMatched: 0,
      propertiesWithoutMatchedUnits: 0,
      unitRecordsProcessed: 0,
      matchedUnitRecords: 0,
      unmatchedUnitRecords: 0,
    };

    if (result.ok) {
      var hasMatchIssues =
        summary.propertiesWithoutMatchedUnits > 0 || summary.unmatchedUnitRecords > 0;
      el.classList.add(hasMatchIssues ? 'result-warning' : 'result-success');

      var warningHtml =
        result.warnings && result.warnings.length
          ? '<div class="result-warning"><p><strong>Notes</strong></p><ul>' +
            result.warnings
              .map(function (w) {
                return '<li>' + escapeHtml(w) + '</li>';
              })
              .join('') +
            '</ul></div>'
          : '';

      el.innerHTML =
        '<p class="result-title">Schedule generated successfully</p>' +
        '<ul class="result-stats">' +
        '<li>Properties written: <strong>' +
        formatCount(summary.propertiesProcessed) +
        '</strong></li>' +
        '<li>Properties with matched units: <strong>' +
        formatCount(summary.propertiesMatched || 0) +
        '</strong></li>' +
        '<li>Properties without matched units: <strong>' +
        formatCount(summary.propertiesWithoutMatchedUnits || 0) +
        '</strong></li>' +
        '<li>Unit records written: <strong>' +
        formatCount(summary.matchedUnitRecords) +
        '</strong></li>' +
        '<li>Unmatched unit records: <strong>' +
        formatCount(summary.unmatchedUnitRecords) +
        '</strong></li>' +
        '</ul>' +
        (result.filename
          ? '<p class="generate-hint">Downloaded ' + escapeHtml(result.filename) + '</p>'
          : '') +
        warningHtml;
      return;
    }

    el.classList.add('result-error');
    var errorHtml =
      result.errors && result.errors.length
        ? '<ul class="result-errors">' +
          result.errors
            .map(function (e) {
              return '<li>' + escapeHtml(e) + '</li>';
            })
            .join('') +
          '</ul>'
        : '<p>Generation could not be completed due to a technical error.</p>';

    el.innerHTML =
      '<p class="result-title">Generation could not complete</p>' + errorHtml;
  }

  function setupGenerateButton() {
    var btn = document.getElementById('generate-btn');
    btn.addEventListener('click', function () {
      if (btn.disabled) return;

      clearGenerateResult();

      try {
        var result = generateStandardisedSchedule(
          APP_STATE.property,
          APP_STATE.unit,
          APP_STATE.mappingConfig
        );
        showGenerateResult(result);
      } catch (err) {
        console.error('Generate failed:', err);
        showGenerateResult({
          ok: false,
          errors: ['An unexpected error occurred while generating the schedule.'],
          warnings: [],
          summary: {
            propertiesProcessed: (APP_STATE.property && APP_STATE.property.dataRowCount) || 0,
            unitRecordsProcessed: (APP_STATE.unit && APP_STATE.unit.dataRowCount) || 0,
            matchedUnitRecords: 0,
            unmatchedUnitRecords: 0,
          },
        });
      }
    });
  }

  function refreshStatus() {
    var propertyStatus = document.getElementById('property-status');
    var unitStatus = document.getElementById('unit-status');
    var mappingStatus = document.getElementById('mapping-status');
    var generateBtn = document.getElementById('generate-btn');
    var generateHint = document.getElementById('generate-hint');

    setStatusLines(
      propertyStatus,
      getFileStatusLines(APP_STATE.property, 'property', 'properties')
    );
    setStatusLines(unitStatus, getFileStatusLines(APP_STATE.unit, 'record', 'records'));

    APP_STATE.mappingConfigured = isMappingConfigured();
    setStatusLines(mappingStatus, getMappingStatusLines());

    // Incomplete matching must NOT disable Generate - only require both files parsed.
    var canGenerate = Boolean(APP_STATE.property && APP_STATE.unit);
    generateBtn.disabled = !canGenerate;

    if (canGenerate) {
      var analysis = analyseMatching(
        APP_STATE.property,
        APP_STATE.unit,
        APP_STATE.mappingConfig
      );
      renderMatchingSummaryPanel(analysis);
      generateHint.textContent =
        'Ready to generate. Incomplete matching will not block the schedule - unmatched records are reported separately.';
    } else {
      renderMatchingSummaryPanel(null);
      generateHint.textContent =
        'Upload Property Level and Unit / TS Level .xlsx files to enable generation.';
    }
  }


  // ---------------------------------------------------------------------------
  // TS Compare ? state, matching, rendering, Excel export
  // Reads Unit mappings from APP_STATE.mappingConfig; never modifies them.
  // ---------------------------------------------------------------------------

  var COMPARE_IDENTITY = {
    property: { section: 'General', field: 'Property Name' },
    unit: { section: 'General', field: 'Suite' },
  };

  var COMPARE_NUMERIC_EPS = 1e-9;
  var COMPARE_OUTPUT_FILENAME = 'Tenancy_Schedule_Comparison.xlsx';

  var COMPARE_STATE = {
    current: null,
    previous: null,
    result: null,
    filter: 'changed',
    search: '',
  };

  var DATE_COMPARE_IDS = {
    start_date: true,
    review_date: true,
    expiry_date: true,
    break_date: true,
  };

  function setAppMode(mode) {
    var app = document.getElementById('app');
    var stdView = document.getElementById('view-standardiser');
    var cmpView = document.getElementById('view-compare');
    var title = document.getElementById('app-title');
    var subtitle = document.getElementById('app-subtitle');
    var stdBtn = document.getElementById('mode-standardiser-btn');
    var cmpBtn = document.getElementById('mode-compare-btn');
    var isCompare = mode === 'compare';

    app.classList.toggle('app-mode-compare', isCompare);
    app.classList.toggle('app-mode-standardiser', !isCompare);
    stdView.classList.toggle('hidden', isCompare);
    cmpView.classList.toggle('hidden', !isCompare);
    stdBtn.classList.toggle('active', !isCompare);
    cmpBtn.classList.toggle('active', isCompare);

    if (isCompare) {
      title.textContent = 'Tenancy Schedule Compare';
      subtitle.textContent =
        'Compare a current tenancy schedule against a previous version and identify unit-level changes.';
      document.title = 'Tenancy Schedule Compare';
    } else {
      title.textContent = 'Tenancy Schedule Standardiser';
      subtitle.textContent =
        'Convert Property and Unit level exports into the standard tenancy schedule.';
      document.title = 'Tenancy Schedule Standardiser';
    }
  }

  function setupAppModeSwitcher() {
    document.getElementById('mode-standardiser-btn').addEventListener('click', function () {
      setAppMode('standardiser');
    });
    document.getElementById('mode-compare-btn').addEventListener('click', function () {
      setAppMode('compare');
    });
    setAppMode('standardiser');
  }

  function getUnitMappedCompareColumns(mappingConfig) {
    var cols = [];
    if (!mappingConfig || !mappingConfig.columnMappings) return cols;
    for (var i = 0; i < STANDARD_TEMPLATE_COLUMNS.length; i += 1) {
      var templateCol = STANDARD_TEMPLATE_COLUMNS[i];
      var entry = null;
      for (var j = 0; j < mappingConfig.columnMappings.length; j += 1) {
        if (mappingConfig.columnMappings[j].targetColumnId === templateCol.id) {
          entry = mappingConfig.columnMappings[j];
          break;
        }
      }
      if (entry && entry.unitMapping && entry.unitMapping.type === 'source') {
        cols.push({
          id: templateCol.id,
          label: templateCol.label,
          unitMapping: entry.unitMapping,
        });
      }
    }
    return cols;
  }

  function isBlankCompareValue(value) {
    if (value === null || value === undefined) return true;
    if (typeof value === 'string' && value.trim() === '') return true;
    return false;
  }

  function excelSerialToUtcDate(serial) {
    // Excel epoch 1899-12-30
    var ms = Math.round((Number(serial) - 25569) * 86400 * 1000);
    var d = new Date(ms);
    return new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
  }

  function toCompareDate(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      return new Date(Date.UTC(value.getFullYear(), value.getMonth(), value.getDate()));
    }
    if (typeof value === 'number' && Number.isFinite(value) && value > 20000 && value < 80000) {
      return excelSerialToUtcDate(value);
    }
    if (typeof value === 'string' && value.trim()) {
      var parsed = Date.parse(value);
      if (!Number.isNaN(parsed)) {
        var d = new Date(parsed);
        return new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
      }
    }
    return null;
  }

  function compareValuesEqual(a, b, columnId) {
    var aBlank = isBlankCompareValue(a);
    var bBlank = isBlankCompareValue(b);
    if (aBlank && bBlank) return true;
    if (aBlank || bBlank) return false;

    if (DATE_COMPARE_IDS[columnId]) {
      var da = toCompareDate(a);
      var db = toCompareDate(b);
      if (da && db) return da.getTime() === db.getTime();
    }

    var aDate = a instanceof Date ? toCompareDate(a) : null;
    var bDate = b instanceof Date ? toCompareDate(b) : null;
    if (aDate && bDate) return aDate.getTime() === bDate.getTime();

    var aNum = typeof a === 'number' && Number.isFinite(a);
    var bNum = typeof b === 'number' && Number.isFinite(b);
    if (aNum && bNum) {
      var diff = Math.abs(a - b);
      var scale = Math.max(1, Math.abs(a), Math.abs(b));
      return diff <= COMPARE_NUMERIC_EPS * scale;
    }

    // Text vs number: allow numeric text equivalence for display-only differences
    if (aNum || bNum) {
      var na = aNum ? a : Number(String(a).replace(/[?$,\s]/g, '').replace(/,/g, ''));
      var nb = bNum ? b : Number(String(b).replace(/[?$,\s]/g, '').replace(/,/g, ''));
      if (Number.isFinite(na) && Number.isFinite(nb)) {
        var d2 = Math.abs(na - nb);
        var s2 = Math.max(1, Math.abs(na), Math.abs(nb));
        return d2 <= COMPARE_NUMERIC_EPS * s2;
      }
    }

    var sa = normaliseJoinValue(a);
    var sb = normaliseJoinValue(b);
    return sa === sb;
  }

  function formatCompareDisplay(value, columnId) {
    if (isBlankCompareValue(value)) return '';
    if (DATE_COMPARE_IDS[columnId] || value instanceof Date) {
      var d = toCompareDate(value);
      if (d) {
        var dd = String(d.getUTCDate()).padStart(2, '0');
        var mm = String(d.getUTCMonth() + 1).padStart(2, '0');
        var yyyy = d.getUTCFullYear();
        return dd + '/' + mm + '/' + yyyy;
      }
    }
    if (typeof value === 'number' && Number.isFinite(value)) {
      if (columnId === 'cap_rate' || columnId.indexOf('pct_') === 0 || columnId === 'niy' || columnId === 'ney' || columnId === 'ry') {
        return (value * 100).toFixed(2) + '%';
      }
      if (Math.abs(value) >= 1000 && Math.floor(value) === value) {
        return value.toLocaleString('en-GB');
      }
      return String(value);
    }
    return String(value);
  }

  function buildUnitCompareRecord(parsedRow, parsedWorkbook, compareColumns, propertyCol, unitCol) {
    var propertyName = propertyCol ? parsedRow.values[propertyCol.key] : null;
    var unitName = unitCol ? parsedRow.values[unitCol.key] : null;
    var values = {};
    for (var i = 0; i < compareColumns.length; i += 1) {
      var col = compareColumns[i];
      values[col.id] = resolveMappedValue(col.unitMapping, parsedRow, parsedWorkbook.columns);
    }
    var propKey = normaliseJoinValue(propertyName);
    var unitKey = normaliseJoinValue(unitName);
    var incompleteIdentity = !propKey || !unitKey;
    // Incomplete identities get a unique key so they never silently match another row.
    var key = incompleteIdentity
      ? '__incomplete__|' + parsedRow.rowIndex + '|' + propKey + '||' + unitKey
      : propKey + '||' + unitKey;
    return {
      sourceRowIndex: parsedRow.rowIndex,
      propertyName: propertyName,
      unitName: unitName,
      key: key,
      incompleteIdentity: incompleteIdentity,
      values: values,
      tenantName: values.tenant_name != null ? values.tenant_name : null,
    };
  }

  function indexCompareRecords(records) {
    var map = new Map();
    var duplicates = new Set();
    for (var i = 0; i < records.length; i += 1) {
      var rec = records[i];
      if (rec.incompleteIdentity) {
        map.set(rec.key, [rec]);
        duplicates.add(rec.key);
        continue;
      }
      var list = map.get(rec.key);
      if (!list) {
        map.set(rec.key, [rec]);
      } else {
        list.push(rec);
        duplicates.add(rec.key);
      }
    }
    return { map: map, duplicates: duplicates };
  }

  function runTenancyCompare(currentParsed, previousParsed, mappingConfig) {
    var errors = [];
    var warnings = [];
    if (!currentParsed) errors.push('Current TS workbook is not loaded.');
    if (!previousParsed) errors.push('Previous TS workbook is not loaded.');
    if (!mappingConfig) errors.push('Mapping configuration is not available.');
    if (errors.length) {
      return { ok: false, errors: errors, warnings: warnings, pairs: [], summary: null, columns: [] };
    }

    var propertyColCurrent = findExactSourceColumn(
      currentParsed.columns,
      COMPARE_IDENTITY.property.section,
      COMPARE_IDENTITY.property.field
    );
    var unitColCurrent = findExactSourceColumn(
      currentParsed.columns,
      COMPARE_IDENTITY.unit.section,
      COMPARE_IDENTITY.unit.field
    );
    var propertyColPrevious = findExactSourceColumn(
      previousParsed.columns,
      COMPARE_IDENTITY.property.section,
      COMPARE_IDENTITY.property.field
    );
    var unitColPrevious = findExactSourceColumn(
      previousParsed.columns,
      COMPARE_IDENTITY.unit.section,
      COMPARE_IDENTITY.unit.field
    );

    if (!propertyColCurrent) errors.push('Property Name field could not be identified in Current TS.');
    if (!unitColCurrent) errors.push('Suite field could not be identified in Current TS.');
    if (!propertyColPrevious) errors.push('Property Name field could not be identified in Previous TS.');
    if (!unitColPrevious) errors.push('Suite field could not be identified in Previous TS.');
    if (errors.length) {
      return { ok: false, errors: errors, warnings: warnings, pairs: [], summary: null, columns: [] };
    }

    var compareColumns = getUnitMappedCompareColumns(mappingConfig);
    if (compareColumns.length === 0) {
      errors.push('No Unit / TS mappings are configured in Advanced Settings, so there are no fields to compare.');
      return { ok: false, errors: errors, warnings: warnings, pairs: [], summary: null, columns: [] };
    }

    var currentRecords = currentParsed.rows.map(function (row) {
      return buildUnitCompareRecord(row, currentParsed, compareColumns, propertyColCurrent, unitColCurrent);
    });
    var previousRecords = previousParsed.rows.map(function (row) {
      return buildUnitCompareRecord(row, previousParsed, compareColumns, propertyColPrevious, unitColPrevious);
    });

    var currentIndex = indexCompareRecords(currentRecords);
    var previousIndex = indexCompareRecords(previousRecords);

    if (currentIndex.duplicates.size > 0) {
      warnings.push(
        currentIndex.duplicates.size +
          ' ambiguous Property Name + Suite key(s) found in Current TS.'
      );
    }
    if (previousIndex.duplicates.size > 0) {
      warnings.push(
        previousIndex.duplicates.size +
          ' ambiguous Property Name + Suite key(s) found in Previous TS.'
      );
    }

    var pairs = [];
    var seenPrev = new Set();
    var propertyOrder = [];
    var propertySeen = {};

    function noteProperty(name) {
      var key = normaliseJoinValue(name) || '(blank property)';
      if (!propertySeen[key]) {
        propertySeen[key] = true;
        propertyOrder.push(name == null || name === '' ? '(Blank Property)' : String(name));
      }
      return key;
    }

    // Walk Current order
    for (var ci = 0; ci < currentRecords.length; ci += 1) {
      var cur = currentRecords[ci];
      noteProperty(cur.propertyName);
      var curList = currentIndex.map.get(cur.key) || [cur];
      var prevList = previousIndex.map.get(cur.key) || [];
      var ambiguous =
        currentIndex.duplicates.has(cur.key) || previousIndex.duplicates.has(cur.key);

      // Only emit once per key from the first current occurrence
      if (curList[0] !== cur) continue;

      if (ambiguous) {
        pairs.push({
          status: 'ambiguous',
          key: cur.key,
          propertyName: cur.propertyName,
          unitName: cur.unitName,
          propertyKey: normaliseJoinValue(cur.propertyName) || '(blank property)',
          changedFields: [],
          changedLabels: [],
          current: cur,
          previous: prevList[0] || null,
        });
        if (prevList[0]) seenPrev.add(cur.key);
        continue;
      }

      if (!prevList.length) {
        pairs.push({
          status: 'new',
          key: cur.key,
          propertyName: cur.propertyName,
          unitName: cur.unitName,
          propertyKey: normaliseJoinValue(cur.propertyName) || '(blank property)',
          changedFields: [],
          changedLabels: ['New unit'],
          current: cur,
          previous: null,
        });
        continue;
      }

      var prev = prevList[0];
      seenPrev.add(cur.key);
      var changedFields = [];
      var changedLabels = [];
      for (var k = 0; k < compareColumns.length; k += 1) {
        var col = compareColumns[k];
        var cv = cur.values[col.id];
        var pv = prev.values[col.id];
        if (!compareValuesEqual(cv, pv, col.id)) {
          changedFields.push(col.id);
          changedLabels.push(col.label);
        }
      }
      pairs.push({
        status: changedFields.length ? 'changed' : 'unchanged',
        key: cur.key,
        propertyName: cur.propertyName,
        unitName: cur.unitName,
        propertyKey: normaliseJoinValue(cur.propertyName) || '(blank property)',
        changedFields: changedFields,
        changedLabels: changedLabels,
        current: cur,
        previous: prev,
      });
    }

    // Removed / leftover previous keys
    for (var pi = 0; pi < previousRecords.length; pi += 1) {
      var prevRec = previousRecords[pi];
      var prevList2 = previousIndex.map.get(prevRec.key) || [prevRec];
      if (prevList2[0] !== prevRec) continue;
      if (seenPrev.has(prevRec.key)) continue;
      if (currentIndex.map.has(prevRec.key)) continue;

      noteProperty(prevRec.propertyName);
      var amb =
        previousIndex.duplicates.has(prevRec.key) || currentIndex.duplicates.has(prevRec.key);
      pairs.push({
        status: amb ? 'ambiguous' : 'removed',
        key: prevRec.key,
        propertyName: prevRec.propertyName,
        unitName: prevRec.unitName,
        propertyKey: normaliseJoinValue(prevRec.propertyName) || '(blank property)',
        changedFields: [],
        changedLabels: amb ? [] : ['Removed unit'],
        current: null,
        previous: prevRec,
      });
    }

    // Sort: property by first-seen order, then current order / removed after
    var propertyRank = {};
    for (var pr = 0; pr < propertyOrder.length; pr += 1) {
      propertyRank[normaliseJoinValue(propertyOrder[pr]) || '(blank property)'] = pr;
    }
    pairs.sort(function (a, b) {
      var ra = propertyRank[a.propertyKey] != null ? propertyRank[a.propertyKey] : 99999;
      var rb = propertyRank[b.propertyKey] != null ? propertyRank[b.propertyKey] : 99999;
      if (ra !== rb) return ra - rb;
      var ua = normaliseJoinValue(a.unitName);
      var ub = normaliseJoinValue(b.unitName);
      if (ua < ub) return -1;
      if (ua > ub) return 1;
      return 0;
    });

    var summary = {
      currentUnits: currentRecords.length,
      previousUnits: previousRecords.length,
      matched: 0,
      changed: 0,
      unchanged: 0,
      newUnits: 0,
      removedUnits: 0,
      ambiguous: 0,
    };
    for (var s = 0; s < pairs.length; s += 1) {
      var st = pairs[s].status;
      if (st === 'changed') {
        summary.changed += 1;
        summary.matched += 1;
      } else if (st === 'unchanged') {
        summary.unchanged += 1;
        summary.matched += 1;
      } else if (st === 'new') summary.newUnits += 1;
      else if (st === 'removed') summary.removedUnits += 1;
      else if (st === 'ambiguous') summary.ambiguous += 1;
    }

    return {
      ok: true,
      errors: [],
      warnings: warnings,
      pairs: pairs,
      summary: summary,
      columns: compareColumns,
      propertyOrder: propertyOrder,
    };
  }

  function setupCompareDropzone(dropzone) {
    var role = dropzone.dataset.compareRole;
    var input = qs(dropzone, '[data-role="input"]');
    var browseBtn = qs(dropzone, '[data-role="browse"]');
    var idle = qs(dropzone, '[data-role="idle"]');
    var success = qs(dropzone, '[data-role="success"]');
    var errorEl = qs(dropzone, '[data-role="error"]');
    var filenameEl = qs(dropzone, '[data-role="filename"]');
    var columnsEl = qs(dropzone, '[data-role="columns"]');
    var rowsEl = qs(dropzone, '[data-role="rows"]');

    function showIdle() {
      idle.classList.remove('hidden');
      success.classList.add('hidden');
      errorEl.classList.add('hidden');
      dropzone.classList.remove('has-file', 'has-error');
    }

    function showError(message) {
      idle.classList.add('hidden');
      success.classList.add('hidden');
      errorEl.classList.remove('hidden');
      errorEl.textContent = message;
      dropzone.classList.remove('has-file');
      dropzone.classList.add('has-error');
    }

    function showSuccess(parsed) {
      idle.classList.add('hidden');
      errorEl.classList.add('hidden');
      success.classList.remove('hidden');
      filenameEl.textContent = '\u2713 ' + parsed.fileName;
      columnsEl.textContent =
        '\u2713 ' + formatCount(parsed.columns.length) + ' source columns detected';
      rowsEl.textContent =
        '\u2713 ' + formatCount(parsed.dataRowCount) + ' data rows detected';
      dropzone.classList.add('has-file');
      dropzone.classList.remove('has-error');
    }

    function handleFile(file) {
      if (!file) return;
      parseSourceFile(file, 'unit')
        .then(function (parsed) {
          COMPARE_STATE[role] = parsed;
          COMPARE_STATE.result = null;
          showSuccess(parsed);
          refreshCompareUi();
        })
        .catch(function (err) {
          COMPARE_STATE[role] = null;
          COMPARE_STATE.result = null;
          showError(
            err && err.message
              ? err.message
              : 'The file could not be processed. Please check it is a valid .xlsx export.'
          );
          refreshCompareUi();
          console.error('[compare:' + role + '] parse failed:', err);
        });
    }

    function openPicker(event) {
      if (event) event.stopPropagation();
      if (!input) return;
      try {
        input.click();
      } catch (err) {
        console.error('File picker could not be opened:', err);
      }
    }

    if (browseBtn) {
      browseBtn.addEventListener('click', function (event) {
        event.stopPropagation();
        openPicker(event);
      });
    }

    dropzone.addEventListener('click', function (event) {
      if (browseBtn && (event.target === browseBtn || browseBtn.contains(event.target))) return;
      if (event.target && event.target.tagName === 'INPUT') return;
      openPicker(event);
    });

    dropzone.addEventListener('keydown', function (event) {
      if (event.key === 'Enter' || event.key === ' ') openPicker(event);
    });

    input.addEventListener('change', function () {
      var file = input.files && input.files[0];
      input.value = '';
      handleFile(file);
    });

    dropzone.addEventListener('dragenter', function (event) {
      event.preventDefault();
      dropzone.classList.add('drag-over');
    });
    dropzone.addEventListener('dragover', function (event) {
      event.preventDefault();
      dropzone.classList.add('drag-over');
    });
    dropzone.addEventListener('dragleave', function (event) {
      if (!dropzone.contains(event.relatedTarget)) dropzone.classList.remove('drag-over');
    });
    dropzone.addEventListener('drop', function (event) {
      event.preventDefault();
      dropzone.classList.remove('drag-over');
      var file = event.dataTransfer && event.dataTransfer.files && event.dataTransfer.files[0];
      handleFile(file);
    });

    showIdle();
  }

  function refreshCompareUi() {
    var compareBtn = document.getElementById('compare-btn');
    var downloadBtn = document.getElementById('compare-download-btn');
    var hint = document.getElementById('compare-hint');
    var errEl = document.getElementById('compare-error');
    var canCompare = Boolean(COMPARE_STATE.current && COMPARE_STATE.previous);
    compareBtn.disabled = !canCompare;
    downloadBtn.disabled = !(COMPARE_STATE.result && COMPARE_STATE.result.ok);

    if (!COMPARE_STATE.current && !COMPARE_STATE.previous) {
      hint.textContent = 'Upload Current and Previous Unit / TS .xlsx files to compare.';
    } else if (!canCompare) {
      hint.textContent = 'Upload both Current and Previous TS files to enable comparison.';
    } else if (!COMPARE_STATE.result) {
      hint.textContent =
        'Ready to compare using Unit / TS mappings from Advanced Settings.';
    }

    if (!COMPARE_STATE.result) {
      document.getElementById('compare-results').classList.add('hidden');
      errEl.classList.add('hidden');
      errEl.textContent = '';
    }
  }

  function renderCompareSummary(summary) {
    var el = document.getElementById('compare-summary');
    var cards = [
      ['Current TS', summary.currentUnits, ''],
      ['Previous TS', summary.previousUnits, ''],
      ['Matched', summary.matched, ''],
      ['Changed', summary.changed, 'stat-changed'],
      ['Unchanged', summary.unchanged, ''],
      ['New Units', summary.newUnits, 'stat-new'],
      ['Removed Units', summary.removedUnits, 'stat-removed'],
      ['Ambiguous', summary.ambiguous, 'stat-ambiguous'],
    ];
    el.innerHTML = cards
      .map(function (c) {
        return (
          '<div class="compare-stat ' +
          c[2] +
          '"><span class="label">' +
          escapeHtml(c[0]) +
          '</span><span class="value">' +
          formatCount(c[1]) +
          '</span></div>'
        );
      })
      .join('');
  }

  function pairMatchesFilter(pair, filter, search) {
    if (filter && filter !== 'all') {
      if (filter === 'changed' && pair.status !== 'changed') return false;
      if (filter === 'unchanged' && pair.status !== 'unchanged') return false;
      if (filter === 'new' && pair.status !== 'new') return false;
      if (filter === 'removed' && pair.status !== 'removed') return false;
      if (filter === 'ambiguous' && pair.status !== 'ambiguous') return false;
    }
    if (search) {
      var q = normaliseJoinValue(search);
      var hay = [
        pair.propertyName,
        pair.unitName,
        pair.current && pair.current.tenantName,
        pair.previous && pair.previous.tenantName,
        pair.changedLabels && pair.changedLabels.join(' '),
      ]
        .map(function (v) {
          return normaliseJoinValue(v);
        })
        .join(' ');
      if (hay.indexOf(q) === -1) return false;
    }
    return true;
  }

  function statusBadge(status) {
    var label = status.toUpperCase();
    return '<span class="badge badge-' + status + '">' + escapeHtml(label) + '</span>';
  }

  function renderCompareTable() {
    var result = COMPARE_STATE.result;
    if (!result || !result.ok) return;

    var head = document.getElementById('compare-table-head');
    var body = document.getElementById('compare-table-body');
    var cols = result.columns;

    var stickyHeaders =
      '<th class="sticky-col sticky-0">Status</th>' +
      '<th class="sticky-col sticky-1">Property</th>' +
      '<th class="sticky-col sticky-2">Unit</th>' +
      '<th class="sticky-col sticky-3">Version</th>' +
      '<th>Changes</th>';
    for (var i = 0; i < cols.length; i += 1) {
      stickyHeaders += '<th>' + escapeHtml(cols[i].label) + '</th>';
    }
    head.innerHTML = '<tr>' + stickyHeaders + '</tr>';

    var filter = COMPARE_STATE.filter || 'changed';
    var search = COMPARE_STATE.search || '';
    var htmlRows = '';
    var lastPropertyKey = null;

    for (var p = 0; p < result.pairs.length; p += 1) {
      var pair = result.pairs[p];
      if (!pairMatchesFilter(pair, filter, search)) continue;

      if (pair.propertyKey !== lastPropertyKey) {
        lastPropertyKey = pair.propertyKey;
        var pname =
          pair.propertyName == null || pair.propertyName === ''
            ? '(Blank Property)'
            : String(pair.propertyName);
        htmlRows +=
          '<tr class="compare-property-band"><td colspan="' +
          (5 + cols.length) +
          '">' +
          escapeHtml(pname) +
          '</td></tr>';
      }

      var changesText =
        pair.status === 'changed'
          ? pair.changedLabels.join(', ')
          : pair.status === 'new'
            ? 'New unit'
            : pair.status === 'removed'
              ? 'Removed unit'
              : pair.status === 'ambiguous'
                ? 'Ambiguous match key'
                : '';

      function rowHtml(version, record, isLastOfPair) {
        var trClass =
          'version-' +
          version +
          (pair.status === 'new' || pair.status === 'removed' || pair.status === 'ambiguous'
            ? ' status-' + pair.status
            : '') +
          (isLastOfPair ? ' pair-sep' : '');
        var versionBadge =
          version === 'current'
            ? '<span class="badge badge-current">CURRENT</span>'
            : '<span class="badge badge-previous">PREVIOUS</span>';

        var cells =
          '<td class="sticky-col sticky-0">' +
          (version === 'current' ? statusBadge(pair.status) : '') +
          '</td>' +
          '<td class="sticky-col sticky-1">' +
          escapeHtml(pair.propertyName == null ? '' : String(pair.propertyName)) +
          '</td>' +
          '<td class="sticky-col sticky-2">' +
          escapeHtml(pair.unitName == null ? '' : String(pair.unitName)) +
          '</td>' +
          '<td class="sticky-col sticky-3">' +
          versionBadge +
          '</td>' +
          '<td>' +
          (version === 'current' ? escapeHtml(changesText) : '') +
          '</td>';

        for (var c = 0; c < cols.length; c += 1) {
          var col = cols[c];
          var changed =
            pair.status === 'changed' && pair.changedFields.indexOf(col.id) !== -1;
          var raw = record ? record.values[col.id] : null;
          var display;
          if (!record) {
            display =
              '<span class="compare-empty">' +
              (version === 'current'
                ? 'Not present in current TS'
                : 'Not present in previous TS') +
              '</span>';
          } else {
            display = escapeHtml(formatCompareDisplay(raw, col.id));
          }
          cells +=
            '<td class="' +
            (changed ? 'cell-changed' : '') +
            '">' +
            display +
            '</td>';
        }
        return '<tr class="' + trClass + '">' + cells + '</tr>';
      }

      htmlRows += rowHtml('current', pair.current, false);
      htmlRows += rowHtml('previous', pair.previous, true);
    }

    body.innerHTML =
      htmlRows ||
      '<tr><td colspan="' +
        (5 + cols.length) +
        '" class="compare-empty">No units match the current filter/search.</td></tr>';
  }

  function showCompareResult(result) {
    var errEl = document.getElementById('compare-error');
    var resultsEl = document.getElementById('compare-results');
    if (!result.ok) {
      resultsEl.classList.add('hidden');
      errEl.classList.remove('hidden');
      errEl.textContent = (result.errors || []).join(' ');
      document.getElementById('compare-download-btn').disabled = true;
      return;
    }
    errEl.classList.add('hidden');
    errEl.textContent = '';
    resultsEl.classList.remove('hidden');
    renderCompareSummary(result.summary);
    renderCompareTable();
    document.getElementById('compare-download-btn').disabled = false;
    if (result.warnings && result.warnings.length) {
      document.getElementById('compare-hint').textContent = result.warnings.join(' ');
    } else {
      document.getElementById('compare-hint').textContent =
        'Comparison complete. Default filter shows Changed units.';
    }
  }

  function downloadCompareWorkbook(result) {
    if (!result || !result.ok) return;
    var cols = result.columns;
    var headers = ['Status', 'Property', 'Unit', 'Version', 'Changes'].concat(
      cols.map(function (c) {
        return c.label;
      })
    );

    var aoa = [headers];
    var metaRows = []; // track styling: {r, status, version, changedFields}

    for (var i = 0; i < result.pairs.length; i += 1) {
      var pair = result.pairs[i];
      var changesText =
        pair.status === 'changed'
          ? pair.changedLabels.join('; ')
          : pair.status === 'new'
            ? 'New unit'
            : pair.status === 'removed'
              ? 'Removed unit'
              : pair.status === 'ambiguous'
                ? 'Ambiguous match key'
                : '';

      function pushVersion(version, record) {
        var row = [
          version === 'current' ? pair.status.toUpperCase() : '',
          pair.propertyName,
          pair.unitName,
          version.toUpperCase(),
          version === 'current' ? changesText : '',
        ];
        for (var c = 0; c < cols.length; c += 1) {
          row.push(record ? record.values[cols[c].id] : null);
        }
        aoa.push(row);
        metaRows.push({
          r: aoa.length - 1,
          status: pair.status,
          version: version,
          changedFields: pair.changedFields || [],
        });
      }
      pushVersion('current', pair.current);
      pushVersion('previous', pair.previous);
    }

    var wb = XLSX.utils.book_new();
    var ws = XLSX.utils.aoa_to_sheet(aoa);

    // Style header + rows
    var range = XLSX.utils.decode_range(ws['!ref']);
    for (var c = range.s.c; c <= range.e.c; c += 1) {
      var addr = XLSX.utils.encode_cell({ r: 0, c: c });
      if (!ws[addr]) continue;
      ws[addr].s = {
        fill: { patternType: 'solid', fgColor: { rgb: '002060' } },
        font: { bold: true, color: { rgb: 'FFFFFF' }, name: 'Calibri', sz: 11 },
        alignment: { vertical: 'center', wrapText: true },
      };
    }

    for (var m = 0; m < metaRows.length; m += 1) {
      var meta = metaRows[m];
      var fill =
        meta.status === 'new'
          ? 'F3FAF6'
          : meta.status === 'removed'
            ? 'FFF7F6'
            : meta.version === 'current'
              ? 'F7F9FC'
              : 'FFFFFF';
      for (var cc = range.s.c; cc <= range.e.c; cc += 1) {
        var cellAddr = XLSX.utils.encode_cell({ r: meta.r, c: cc });
        if (!ws[cellAddr]) ws[cellAddr] = { t: 's', v: '' };
        var colId = cc >= 5 ? cols[cc - 5].id : null;
        var changed =
          meta.status === 'changed' && colId && meta.changedFields.indexOf(colId) !== -1;
        ws[cellAddr].s = {
          fill: {
            patternType: 'solid',
            fgColor: { rgb: changed ? 'FFF3BF' : fill },
          },
          font: { name: 'Calibri', sz: 10, bold: meta.version === 'current' },
        };
        if (colId && DATE_COMPARE_IDS[colId] && ws[cellAddr].v instanceof Date) {
          ws[cellAddr].t = 'd';
          ws[cellAddr].z = 'dd/mm/yyyy';
          ws[cellAddr].s.numFmt = 'dd/mm/yyyy';
        } else if (colId && typeof ws[cellAddr].v === 'number') {
          var fmt = getNumberFormat(colId, ws[cellAddr].v);
          if (fmt) {
            ws[cellAddr].z = fmt;
            ws[cellAddr].s.numFmt = fmt;
          }
        }
      }
    }

    ws['!views'] = [{ state: 'frozen', ySplit: 1, topLeftCell: 'A2', activeCell: 'A2' }];
    ws['!autofilter'] = { ref: ws['!ref'] };
    ws['!cols'] = headers.map(function (h) {
      return { wch: Math.min(28, Math.max(10, String(h).length + 2)) };
    });
    XLSX.utils.book_append_sheet(wb, ws, 'TS Comparison');

    // Change Summary sheet
    var summaryAoa = [['Status', 'Property', 'Unit', 'Changed Fields']];
    for (var si = 0; si < result.pairs.length; si += 1) {
      var sp = result.pairs[si];
      summaryAoa.push([
        sp.status.toUpperCase(),
        sp.propertyName,
        sp.unitName,
        sp.changedLabels.join('; '),
      ]);
    }
    var ws2 = XLSX.utils.aoa_to_sheet(summaryAoa);
    for (var sc = 0; sc < 4; sc += 1) {
      var saddr = XLSX.utils.encode_cell({ r: 0, c: sc });
      if (ws2[saddr]) {
        ws2[saddr].s = {
          fill: { patternType: 'solid', fgColor: { rgb: '002060' } },
          font: { bold: true, color: { rgb: 'FFFFFF' } },
        };
      }
    }
    ws2['!views'] = [{ state: 'frozen', ySplit: 1, topLeftCell: 'A2', activeCell: 'A2' }];
    ws2['!autofilter'] = { ref: ws2['!ref'] };
    XLSX.utils.book_append_sheet(wb, ws2, 'Change Summary');

    downloadWorkbook(wb, COMPARE_OUTPUT_FILENAME);
  }

  function clearComparison() {
    COMPARE_STATE.current = null;
    COMPARE_STATE.previous = null;
    COMPARE_STATE.result = null;
    COMPARE_STATE.filter = 'changed';
    COMPARE_STATE.search = '';

    document.querySelectorAll('#view-compare .dropzone').forEach(function (dz) {
      dz.classList.remove('has-file', 'has-error', 'drag-over');
      var idle = dz.querySelector('[data-role="idle"]');
      var success = dz.querySelector('[data-role="success"]');
      var error = dz.querySelector('[data-role="error"]');
      if (idle) idle.classList.remove('hidden');
      if (success) success.classList.add('hidden');
      if (error) {
        error.classList.add('hidden');
        error.textContent = '';
      }
      var input = dz.querySelector('[data-role="input"]');
      if (input) input.value = '';
    });

    document.getElementById('compare-results').classList.add('hidden');
    document.getElementById('compare-error').classList.add('hidden');
    document.getElementById('compare-error').textContent = '';
    document.getElementById('compare-search').value = '';
    document.querySelectorAll('.compare-filter-btn').forEach(function (btn) {
      btn.classList.toggle('active', btn.dataset.compareFilter === 'changed');
    });
    refreshCompareUi();
  }

  function setupCompareTool() {
    document.querySelectorAll('#view-compare .dropzone').forEach(setupCompareDropzone);

    document.getElementById('compare-btn').addEventListener('click', function () {
      var result = runTenancyCompare(
        COMPARE_STATE.current,
        COMPARE_STATE.previous,
        APP_STATE.mappingConfig
      );
      COMPARE_STATE.result = result;
      showCompareResult(result);
      refreshCompareUi();
    });

    document.getElementById('compare-download-btn').addEventListener('click', function () {
      if (COMPARE_STATE.result && COMPARE_STATE.result.ok) {
        downloadCompareWorkbook(COMPARE_STATE.result);
      }
    });

    document.getElementById('compare-clear-btn').addEventListener('click', clearComparison);

    document.querySelectorAll('.compare-filter-btn').forEach(function (btn) {
      btn.addEventListener('click', function () {
        COMPARE_STATE.filter = btn.dataset.compareFilter;
        document.querySelectorAll('.compare-filter-btn').forEach(function (b) {
          b.classList.toggle('active', b === btn);
        });
        if (COMPARE_STATE.result && COMPARE_STATE.result.ok) renderCompareTable();
      });
    });

    document.getElementById('compare-search').addEventListener('input', function (event) {
      COMPARE_STATE.search = event.target.value || '';
      if (COMPARE_STATE.result && COMPARE_STATE.result.ok) renderCompareTable();
    });

    refreshCompareUi();
  }

  function init() {
    initialiseMappingState();
    document.querySelectorAll('#view-standardiser .dropzone').forEach(setupDropzone);
    setupAppModeSwitcher();
    setupCompareTool();
    setupSettingsUI({ onConfigChanged: refreshStatus });
    setupGenerateButton();
    refreshStatus();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
