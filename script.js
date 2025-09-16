'use strict';

document.addEventListener('DOMContentLoaded', function() {
    // DOM elements
    const excelFile = document.getElementById('excelFile');
    const downloadAllBtn = document.getElementById('downloadAllBtn');
    const clearAllBtn = document.getElementById('clearAllBtn');
    const previewTable = document.getElementById('previewTable');
    const editTable = document.getElementById('editTable');
    const previewTableSection = document.getElementById('previewTableSection');
    const editTableSection = document.getElementById('editTableSection');
    const previewMode = document.getElementById('previewMode');
    const editMode = document.getElementById('editMode');
    const saveChangesBtn = document.getElementById('saveChangesBtn');
    const processingIndicator = document.getElementById('processingIndicator');
    const adjustmentSummary = document.getElementById('adjustmentSummary');
    
    // Debug: Check if elements exist
    console.log('DOM elements found:', {
        excelFile: !!excelFile,
        downloadAllBtn: !!downloadAllBtn,
        clearAllBtn: !!clearAllBtn
    });
    
    // Data storage
    let xmlDataArray = [];
    let barcodeDataArray = [];
    let originalHeaders = [];
    let rawData = [];
    let adjustedRows = new Set();
    
    // Columns to show in edit mode
    const editableColumns = [
        'Wi_Number', 'Limp_ISBN', 'Cased_ISBN', 'Title',
        'Trim_Height', 'Trim_Width', 'Page_Extent', 'Spine_Size',
        'Reel_Width', 'Cut_Off', 'Imposition', 'Paper_Code'
    ];
    
    // Fixed trim off head value (3mm = 0030)
    function getTrimOffHeadValue() {
        return "0030";
    }
    
    // Function to pad a number with leading zeros to a specific length
    function padWithZeros(num, targetLength) {
        return num.toString().padStart(targetLength, '0');
    }
    
    // Function to generate DataMatrix barcode string from Excel data
    function generateBarcodeString(row) {
        try {
            const limpIsbnIndex = originalHeaders.indexOf('Limp_ISBN');
            const casedIsbnIndex = originalHeaders.indexOf('Cased_ISBN');
            const heightIndex = originalHeaders.indexOf('Trim_Height');
            const widthIndex = originalHeaders.indexOf('Trim_Width');
            const spineIndex = originalHeaders.indexOf('Spine_Size');
            const cutOffIndex = originalHeaders.indexOf('Cut_Off');
            
            // Get ISBN and determine transfer station
            let isbn = '';
            let transferStation = '1';
            
            const limpIsbnValue = limpIsbnIndex !== -1 ? row[limpIsbnIndex] : '';
            const casedIsbnValue = casedIsbnIndex !== -1 ? row[casedIsbnIndex] : '';
            
            if (limpIsbnValue && limpIsbnValue.trim() !== '') {
                isbn = limpIsbnValue;
                transferStation = '1';
            } else if (casedIsbnValue && casedIsbnValue.trim() !== '') {
                isbn = casedIsbnValue;
                transferStation = '2';
            } else {
                return null;
            }
            
            // Clean and format ISBN
            isbn = String(isbn).replace(/\D/g, '');
            if (isbn.length > 13) isbn = isbn.substring(0, 13);
            else if (isbn.length < 13) isbn = isbn.padStart(13, '0');
            
            // Get other values
            const height = parseFloat(row[heightIndex]) || 0;
            const width = parseFloat(row[widthIndex]) || 0;
            const spine = parseFloat(row[spineIndex]) || 0;
            const cutOff = parseFloat(row[cutOffIndex]) || 0;
            
            // Format barcode string (37 digits total)
            const endsheetHeight = "0000";
            
            // Spine formatting
            const spineRounded = Math.round(spine);
            let spineSegment;
            if (spineRounded >= 10) {
                spineSegment = spineRounded.toString() + "0";
            } else {
                spineSegment = "0" + spineRounded.toString() + "0";
            }
            
            const bbHeightSegment = padWithZeros(Math.round(cutOff), 3) + "0";
            const trimOffSegment = getTrimOffHeadValue();
            const trimHeightSegment = padWithZeros(Math.round(height), 3) + "0";
            const trimWidthSegment = padWithZeros(Math.round(width), 3) + "0";
            
            const barcodeString = 
                isbn + endsheetHeight + spineSegment + bbHeightSegment + 
                trimOffSegment + trimHeightSegment + trimWidthSegment + transferStation;
            
            if (barcodeString.length !== 37 || /\D/.test(barcodeString)) {
                return null;
            }
            
            return barcodeString;
        } catch (error) {
            console.error("Error generating barcode:", error);
            return null;
        }
    }
    
    // Function to create DataMatrix SVG using the proper library
    function createDataMatrixSVG(barcodeString) {
        if (!barcodeString || barcodeString.length !== 37) return null;
        
        try {
            return DATAMatrix({
                msg: barcodeString,
                dim: 100,
                rct: 1,
                pad: 2,
                pal: ['#000000', '#ffffff']
            });
        } catch (error) {
            console.error('Error creating DataMatrix:', error);
            return null;
        }
    }
    
    function showProcessingIndicator(text) {
        if (processingIndicator) {
            document.getElementById('processingText').textContent = text;
            processingIndicator.style.display = 'block';
        }
    }

    function hideProcessingIndicator() {
        if (processingIndicator) {
            processingIndicator.style.display = 'none';
        }
    }

    function showAdjustmentSummary(count) {
        if (adjustmentSummary) {
            const adjustmentText = document.getElementById('adjustmentText');
            adjustmentText.textContent = `${count} row(s) with "Limp P/Bound 8pp Cover" production route detected. Trim_Width values have been automatically increased by 10.`;
            adjustmentSummary.classList.remove('d-none');
        }
    }

    function hideAdjustmentSummary() {
        if (adjustmentSummary) {
            adjustmentSummary.classList.add('d-none');
        }
    }

    async function downloadAllFiles() {
        console.log('Download function called');
        showProcessingIndicator('Generating XML files and DataMatrix PDFs...');
        
        try {
            const zip = new JSZip();
            const { jsPDF } = window.jspdf;
            
            // Add all XML files to ZIP
            xmlDataArray.forEach(data => {
                zip.file(`${data.wiNumber}.xml`, data.xml);
            });
            
            // Add all DataMatrix PDFs to ZIP
            for (let i = 0; i < barcodeDataArray.length; i++) {
                const barcodeData = barcodeDataArray[i];
                
                if (!barcodeData.datamatrixSvg || !barcodeData.barcodeString) {
                    console.warn(`Skipping barcode PDF for row ${i}: no valid barcode data`);
                    continue;
                }
                
                try {
                    const svgData = new XMLSerializer().serializeToString(barcodeData.datamatrixSvg);
                    const canvas = document.createElement('canvas');
                    const ctx = canvas.getContext('2d');
                    
                    await new Promise((resolve, reject) => {
                        const img = new Image();
                        img.onload = function() {
                            const widthScaleFactor = 1.10;
                            const heightScaleFactor = 1.27;
                            const barcodeWidth = 17 * widthScaleFactor;
                            const barcodeHeight = 6 * heightScaleFactor;
                            const margin = 10;
                            
                            const pageWidth = barcodeWidth + (margin * 2);
                            const pageHeight = barcodeHeight + (margin * 2);
                            
                            const pdf = new jsPDF({
                                orientation: 'portrait',
                                unit: 'mm',
                                format: [pageHeight, pageWidth]
                            });
                            
                            canvas.width = img.width;
                            canvas.height = img.height;
                            ctx.drawImage(img, 0, 0);
                            
                            const xOffset = (pdf.internal.pageSize.getWidth() - barcodeWidth) / 2;
                            const yOffset = margin;
                            
                            pdf.addImage(
                                canvas.toDataURL('image/png'),
                                'PNG',
                                xOffset,
                                yOffset,
                                barcodeWidth,
                                barcodeHeight
                            );
                            
                            let filenameISBN = barcodeData.isbn;
                            if (barcodeData.barcodeString && barcodeData.barcodeString.length >= 13) {
                                filenameISBN = barcodeData.barcodeString.substring(0, 13);
                            }
                            filenameISBN = filenameISBN.replace(/\D/g, '');
                            
                            const filename = `${barcodeData.wiNumber}_${filenameISBN}_DBC.pdf`;
                            
                            const pdfArrayBuffer = pdf.output('arraybuffer');
                            zip.file(filename, pdfArrayBuffer);
                            
                            resolve();
                        };
                        
                        img.onerror = function() {
                            console.error(`Error loading image for row ${i}`);
                            reject(new Error(`Image load failed for row ${i}`));
                        };
                        
                        img.src = 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(svgData)));
                    });
                    
                } catch (error) {
                    console.error(`Error processing barcode ${i}:`, error);
                }
            }
            
            // Generate and download the combined ZIP
            const content = await zip.generateAsync({ type: "blob" });
            const url = window.URL.createObjectURL(content);
            const a = document.createElement('a');
            a.href = url;
            a.download = "xml_and_datamatrix_files.zip";
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);
            
        } catch (error) {
            console.error('Error generating combined ZIP:', error);
            alert('Error generating files: ' + error.message);
        } finally {
            hideProcessingIndicator();
        }
    }

    function downloadSingleXml(xml, wiNumber) {
        const blob = new Blob([xml], { type: 'application/xml' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${wiNumber}.xml`;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
    }
    
    function downloadSingleBarcodePDF(barcodeData) {
        if (!barcodeData.datamatrixSvg || !barcodeData.barcodeString) {
            alert('No valid barcode data available for this row.');
            return;
        }
        
        try {
            const { jsPDF } = window.jspdf;
            const svgData = new XMLSerializer().serializeToString(barcodeData.datamatrixSvg);
            
            const canvas = document.createElement('canvas');
            const ctx = canvas.getContext('2d');
            
            const img = new Image();
            img.onload = function() {
                const widthScaleFactor = 1.10;
                const heightScaleFactor = 1.27;
                const barcodeWidth = 17 * widthScaleFactor;
                const barcodeHeight = 6 * heightScaleFactor;
                const margin = 10;
                
                const pageWidth = barcodeWidth + (margin * 2);
                const pageHeight = barcodeHeight + (margin * 2);
                
                const pdf = new jsPDF({
                    orientation: pageWidth > pageHeight ? 'landscape' : 'portrait',
                    unit: 'mm',
                    format: [Math.max(pageWidth, pageHeight), Math.min(pageWidth, pageHeight)]
                });
                
                canvas.width = img.width;
                canvas.height = img.height;
                ctx.drawImage(img, 0, 0);
                
                const xOffset = (pdf.internal.pageSize.getWidth() - barcodeWidth) / 2;
                const yOffset = margin;
                
                pdf.addImage(
                    canvas.toDataURL('image/png'),
                    'PNG',
                    xOffset,
                    yOffset,
                    barcodeWidth,
                    barcodeHeight
                );
                
                let filenameISBN = barcodeData.isbn;
                if (barcodeData.barcodeString && barcodeData.barcodeString.length >= 13) {
                    filenameISBN = barcodeData.barcodeString.substring(0, 13);
                }
                filenameISBN = filenameISBN.replace(/\D/g, '');
                
                const filename = `${barcodeData.wiNumber}_${filenameISBN}_DBC.pdf`;
                pdf.save(filename);
            };
            
            img.onerror = function() {
                alert('Error converting SVG to image');
            };
            
            img.src = 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(svgData)));
            
        } catch (error) {
            console.error('PDF generation error:', error);
            alert('Error generating PDF: ' + error.message);
        }
    }

    function deleteRow(index) {
        if (confirm('Are you sure you want to delete this row?')) {
            xmlDataArray.splice(index, 1);
            barcodeDataArray.splice(index, 1);
            rawData.splice(index, 1);
            
            const newAdjustedRows = new Set();
            adjustedRows.forEach(adjustedIndex => {
                if (adjustedIndex < index) {
                    newAdjustedRows.add(adjustedIndex);
                } else if (adjustedIndex > index) {
                    newAdjustedRows.add(adjustedIndex - 1);
                }
            });
            adjustedRows = newAdjustedRows;
            
            updatePreviewTable();
            
            if (editTableSection.style.display !== 'none') {
                populateEditTable();
            }
            
            if (downloadAllBtn) downloadAllBtn.disabled = xmlDataArray.length === 0;
            if (clearAllBtn) clearAllBtn.disabled = xmlDataArray.length === 0;
            
            const rowCountInfo = document.getElementById('rowCountInfo');
            if (rowCountInfo) {
                const rowCount = rawData.length;
                rowCountInfo.textContent = `${rowCount} record${rowCount !== 1 ? 's' : ''} found`;
            }
        }
    }

    function updatePreviewTable() {
        const tbody = previewTable.querySelector('tbody');
        tbody.innerHTML = '';
        
        xmlDataArray.forEach((xmlData, index) => {
            const barcodeData = barcodeDataArray[index];
            const row = document.createElement('tr');
            row.dataset.originalIndex = index;
            
            if (adjustedRows.has(index)) {
                row.classList.add('trim-width-adjusted');
            }
            
            // Delete cell (first column)
            const deleteCell = document.createElement('td');
            deleteCell.className = 'text-center';
            const deleteBtn = document.createElement('button');
            deleteBtn.className = 'btn btn-danger btn-sm';
            deleteBtn.innerHTML = '<i class="bi bi-trash"></i>';
            deleteBtn.title = 'Delete Row';
            deleteBtn.onclick = () => deleteRow(index);
            deleteCell.appendChild(deleteBtn);
            row.appendChild(deleteCell);
            
            // Wi Number cell
            const wiNumberCell = document.createElement('td');
            wiNumberCell.textContent = xmlData.wiNumber;
            row.appendChild(wiNumberCell);
            
            // ISBN cell
            const isbnCell = document.createElement('td');
            isbnCell.textContent = xmlData.isbn;
            row.appendChild(isbnCell);
            
            // Title cell
            const titleCell = document.createElement('td');
            titleCell.textContent = xmlData.title;
            row.appendChild(titleCell);
            
            // Production Route cell
            const productionRouteCell = document.createElement('td');
            productionRouteCell.textContent = xmlData.productionRoute;
            row.appendChild(productionRouteCell);
            
            // Trim Width cell
            const trimWidthCell = document.createElement('td');
            trimWidthCell.textContent = xmlData.trimWidth;
            if (adjustedRows.has(index)) {
                trimWidthCell.innerHTML = `<strong>${xmlData.trimWidth}</strong> <i class="bi bi-plus-circle-fill text-warning" title="Adjusted +10"></i>`;
            }
            row.appendChild(trimWidthCell);
            
            // DataMatrix cell
            const dataMatrixCell = document.createElement('td');
            dataMatrixCell.className = 'datamatrix-cell';
            if (barcodeData && barcodeData.datamatrixSvg) {
                const svgClone = barcodeData.datamatrixSvg.cloneNode(true);
                svgClone.classList.add('datamatrix-preview');
                dataMatrixCell.appendChild(svgClone);
                
                svgClone.style.cursor = 'pointer';
                svgClone.title = 'Click to download DataMatrix PDF';
                svgClone.addEventListener('click', () => downloadSingleBarcodePDF(barcodeData));
            } else {
                dataMatrixCell.textContent = 'N/A';
                dataMatrixCell.style.color = '#6c757d';
            }
            row.appendChild(dataMatrixCell);
            
            // Action cell (last column)
            const actionCell = document.createElement('td');
            actionCell.className = 'text-center';
            const downloadBtn = document.createElement('button');
            downloadBtn.className = 'btn btn-primary btn-sm';
            downloadBtn.innerHTML = '<i class="bi bi-download"></i>';
            downloadBtn.title = 'Download XML';
            downloadBtn.onclick = () => downloadSingleXml(xmlData.xml, xmlData.wiNumber);
            actionCell.appendChild(downloadBtn);
            row.appendChild(actionCell);
            
            tbody.appendChild(row);
        });
    }

    function regenerateAllData() {
        showProcessingIndicator('Regenerating XML and barcode data...');
        
        xmlDataArray = [];
        barcodeDataArray = [];
        
        for (let i = 0; i < rawData.length; i++) {
            const values = rawData[i];
            if (values.length === 0) continue;
            
            const xml = createXmlFromRow(originalHeaders, values);
            const barcodeString = generateBarcodeString(values);
            const datamatrixSvg = barcodeString ? createDataMatrixSVG(barcodeString) : null;
            
            const wiNumberIndex = originalHeaders.indexOf('Wi_Number');
            const limpIsbnIndex = originalHeaders.indexOf('Limp_ISBN');
            const casedIsbnIndex = originalHeaders.indexOf('Cased_ISBN');
            const titleIndex = originalHeaders.indexOf('Title');
            const productionRouteIndex = originalHeaders.indexOf('Production_Route');
            const trimWidthIndex = originalHeaders.indexOf('Trim_Width');
            
            const wiNumber = values[wiNumberIndex] || 'unknown';
            const limpIsbn = values[limpIsbnIndex] || '';
            const casedIsbn = values[casedIsbnIndex] || '';
            const isbn = limpIsbn || casedIsbn || 'N/A';
            
            let title = values[titleIndex] || 'N/A';
            if (title !== 'N/A') {
                title = title.toString().replace(/,/g, '');
            }
            
            const productionRoute = values[productionRouteIndex] || 'N/A';
            const trimWidth = values[trimWidthIndex] || 'N/A';
            
            xmlDataArray.push({
                wiNumber: wiNumber,
                isbn: isbn,
                title: title,
                productionRoute: productionRoute,
                trimWidth: trimWidth,
                xml: xml
            });
            
            barcodeDataArray.push({
                wiNumber: wiNumber,
                isbn: isbn,
                barcodeString: barcodeString,
                datamatrixSvg: datamatrixSvg,
                rowIndex: i
            });
        }
        
        hideProcessingIndicator();
    }

    function createXmlFromRow(headers, values) {
        const xmlDoc = document.implementation.createDocument(null, "csv", null);
        const root = xmlDoc.documentElement;
        
        const pi = xmlDoc.createProcessingInstruction('xml', 'version="1.0" encoding="UTF-8"');
        xmlDoc.insertBefore(pi, root);
        
        for (let i = 0; i < headers.length; i++) {
            if (headers[i]) {
                const elementName = headers[i].replace(/\s+/g, '_');
                const element = xmlDoc.createElement(elementName);
                
                let value = values[i] !== undefined ? values[i].toString().trim() : '';
                
                if (elementName === 'Title') {
                    value = value.replace(/,/g, '');
                }
                
                element.textContent = value;
                root.appendChild(element);
            }
        }
        
        const serializer = new XMLSerializer();
        let xmlString = serializer.serializeToString(xmlDoc);
        
        xmlString = xmlString.replace(/></g, '>\n<');
        xmlString = xmlString.replace(/<csv>/, '<csv>\n');
        xmlString = xmlString.replace(/<\/csv>/, '\n</csv>');
        xmlString = xmlString.replace(/<([^/?][^>]*)>/g, '  <$1>');
        
        return xmlString;
    }

    function applyProductionRouteLogic() {
        showProcessingIndicator('Applying production route logic...');
        
        const productionRouteIndex = originalHeaders.indexOf('Production_Route');
        const trimWidthIndex = originalHeaders.indexOf('Trim_Width');
        
        if (productionRouteIndex === -1 || trimWidthIndex === -1) {
            hideProcessingIndicator();
            return 0;
        }
        
        let adjustmentCount = 0;
        adjustedRows.clear();
        
        for (let i = 0; i < rawData.length; i++) {
            const row = rawData[i];
            const productionRoute = row[productionRouteIndex];
            const currentTrimWidth = row[trimWidthIndex];
            
            if (productionRoute && productionRoute.toString().trim() === 'Limp P/Bound 8pp Cover') {
                let trimWidthValue = parseFloat(currentTrimWidth) || 0;
                trimWidthValue += 10;
                rawData[i][trimWidthIndex] = trimWidthValue.toString();
                adjustedRows.add(i);
                adjustmentCount++;
            }
        }
        
        hideProcessingIndicator();
        
        if (adjustmentCount > 0) {
            showAdjustmentSummary(adjustmentCount);
        } else {
            hideAdjustmentSummary();
        }
        
        return adjustmentCount;
    }

    function convertExcelToXml(file) {
        showProcessingIndicator('Reading Excel file...');
        
        const reader = new FileReader();
        
        reader.onload = function(e) {
            showProcessingIndicator('Parsing Excel data...');
            
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {
                type: 'array',
                cellDates: true,
                cellNF: true,
                cellText: true,
                raw: true
            });

            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false });

            originalHeaders = jsonData[0];
            rawData = jsonData.slice(1).filter(row => row.length > 0);
            
            const adjustmentCount = applyProductionRouteLogic();
            
            showProcessingIndicator('Generating XML and DataMatrix data...');
            
            xmlDataArray = [];
            barcodeDataArray = [];
            regenerateAllData();
            
            showProcessingIndicator('Updating interface...');
            
            updatePreviewTable();
            
            if (downloadAllBtn) downloadAllBtn.disabled = xmlDataArray.length === 0;
            if (clearAllBtn) clearAllBtn.disabled = xmlDataArray.length === 0;
            
            hideProcessingIndicator();
            
            console.log(`Processing complete. ${adjustmentCount} rows adjusted for Production Route logic.`);
        };

        reader.readAsArrayBuffer(file);
    }

    function clearAll() {
        if (excelFile) excelFile.value = '';
        if (downloadAllBtn) downloadAllBtn.disabled = true;
        if (clearAllBtn) clearAllBtn.disabled = true;
        if (previewTable) previewTable.querySelector('tbody').innerHTML = '';
        xmlDataArray = [];
        barcodeDataArray = [];
        originalHeaders = [];
        rawData = [];
        adjustedRows.clear();
        
        hideProcessingIndicator();
        hideAdjustmentSummary();
        
        if (previewMode) previewMode.checked = true;
        if (previewTableSection) previewTableSection.style.display = 'block';
        if (editTableSection) editTableSection.style.display = 'none';
    }

    // Mode Toggle Buttons
    if (previewMode) {
        previewMode.addEventListener('change', function() {
            if(this.checked) {
                previewTableSection.style.display = 'block';
                editTableSection.style.display = 'none';
            }
        });
    }
    
    if (editMode) {
        editMode.addEventListener('change', function() {
            if(this.checked) {
                previewTableSection.style.display = 'none';
                editTableSection.style.display = 'block';
            }
        });
    }

    // Event Listeners
    if (excelFile) {
        excelFile.addEventListener('change', function() {
            const file = this.files[0];
            if (file) {
                convertExcelToXml(file);
            } else {
                clearAll();
            }
        });
    }

    if (clearAllBtn) {
        clearAllBtn.addEventListener('click', clearAll);
    }
    
    if (downloadAllBtn) {
        downloadAllBtn.addEventListener('click', downloadAllFiles);
    }
    
    console.log('Script loaded successfully');
});