function createPlantingPlan(orders, orderDate) {
    // Get headers from the sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ordersSheet = ss.getSheetByName('Standing Orders');
    const headers = ordersSheet.getRange(1, 1, 1, ordersSheet.getLastColumn()).getValues()[0];
    
    // Get column indices
    const productIndex = headers.indexOf('Product');
    const quantityIndex = headers.indexOf('Quantity Ordered');
    const typeIndex = headers.indexOf('Quantity Type');

    // First, calculate total grams needed for each mix
    const mixTotals = orders.reduce((acc, row) => {
        const product = row[productIndex]?.toString().trim();
        const quantity = row[quantityIndex];
        const packagingType = row[typeIndex]?.toString().trim();

        if (product && quantity && packagingType) {
            if (!acc[product]) {
                acc[product] = 0;
            }
            const gramsForOrder = getGramsPerPackaging(product, packagingType) * quantity;
            acc[product] += gramsForOrder;
        }
        return acc;
    }, {});

    // Then break down mixes into individual varieties
    const varietyGrams = calculateGramsForVarieties(mixTotals);

    // Generate HTML for planting plan using variety breakdown
    const plantingSectionsHtml = generatePlantingSections(varietyGrams);

    // Create template and set values
    const template = HtmlService.createTemplateFromFile('PlantingPlanTemplate');
    template.date = new Date(orderDate).toLocaleDateString();
    template.plantingSections = plantingSectionsHtml || '';

    // Create PDF
    const html = template.evaluate().getContent();
    const blob = Utilities.newBlob(html, 'text/html', 'temp.html');
    const pdf = blob.getAs('application/pdf');
    pdf.setName(`Planting_Plan_${orderDate}.pdf`);
    
    return pdf;
}

function calculateGramsForVarieties(orders) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mixSheet = ss.getSheetByName('Mix Ratios');
    const mixData = mixSheet.getDataRange().getValues();
    
    let gramsForVariety = {};
    
    // For each mix in the orders
    for (let mixName in orders) {
        // Find the mix in the Mix Ratios sheet
        for (let i = 1; i < mixData.length; i++) {
            if (mixData[i][0] === mixName) {  // If mix name matches
                // Get the varieties and their percentages
                const variety = mixData[i][1];
                const percentage = mixData[i][2];
                
                // Initialize if not exists
                if (!gramsForVariety[variety]) {
                    gramsForVariety[variety] = 0;
                }
                
                // Check if percentage is valid and greater than zero
                if (typeof percentage === 'number' && percentage > 0) {
                    // Calculate grams for this variety based on the mix total
                    let grams = (orders[mixName] || 0) * (percentage / 100);
                    gramsForVariety[variety] += grams;
                }
            }
        }
    }
    
    return gramsForVariety;
}

function generatePlantingSections(varietyGrams) {
    return Object.entries(varietyGrams)
        .filter(([variety, grams]) => variety && grams > 0)
        .map(([variety, grams]) => `
            <div class="product-section">
                <div class="product-header">${variety}</div>
                <table>
                    <tr>
                        <th>Total Grams Needed</th>
                        <th>Trays Required</th>
                        <th>Seeds Needed (g)</th>
                        <th>Notes</th>
                    </tr>
                    <tr>
                        <td>${grams.toFixed(2)}g</td>
                        <td>${calculateTraysNeeded(variety, grams)}</td>
                        <td>${calculateSeedsNeeded(variety, grams)}</td>
                        <td></td>
                    </tr>
                </table>
            </div>
        `).join('');
}

function calculateTraysNeeded(variety, totalGrams) {
    const ss = SpreadsheetApp.openById('14UbiYfjOonBZY-55RVPujLVKJsGBd43J6L-9q6ZiSOU');
    const metricsSheet = ss.getSheetByName('Yield & Seed Metrics');
    const metricsData = metricsSheet.getDataRange().getValues();
    
    // Find the yield per tray for this variety
    for (let i = 1; i < metricsData.length; i++) {
        if (metricsData[i][0] === variety) {
            const yieldPerTray = metricsData[i][1];  // Column B: Avg Yield/Tray
            if (yieldPerTray > 0) {
                return Math.ceil(totalGrams / yieldPerTray);
            }
        }
    }
    
    console.error(`No yield data found for variety: ${variety}`);
    return 0;
}

function calculateSeedsNeeded(variety, totalGrams) {
    const ss = SpreadsheetApp.openById('14UbiYfjOonBZY-55RVPujLVKJsGBd43J6L-9q6ZiSOU');
    const metricsSheet = ss.getSheetByName('Yield & Seed Metrics');
    const metricsData = metricsSheet.getDataRange().getValues();
    
    // Find the seeding rate for this variety
    for (let i = 1; i < metricsData.length; i++) {
        if (metricsData[i][0] === variety) {
            const seedsPerTray = metricsData[i][2];  // Column C: Seeds/Tray
            const yieldPerTray = metricsData[i][1];  // Column B: Avg Yield/Tray
            if (yieldPerTray > 0 && seedsPerTray > 0) {
                const traysNeeded = Math.ceil(totalGrams / yieldPerTray);
                return (traysNeeded * seedsPerTray).toFixed(2);
            }
        }
    }
    
    console.error(`No seeding rate found for variety: ${variety}`);
    return 0;
}

function getGramsPerPackaging(product, packagingType) {
    // Use the same spreadsheet ID as in plant-schedule.gs
    const externalSpreadsheetId = '14UbiYfjOonBZY-55RVPujLVKJsGBd43J6L-9q6ZiSOU';
    const packagingYieldSheet = SpreadsheetApp.openById(externalSpreadsheetId).getSheetByName('Packaging Yield');
    const yieldData = packagingYieldSheet.getDataRange().getValues();
    
    for (let i = 1; i < yieldData.length; i++) {
        if (yieldData[i][0] === product && yieldData[i][1] === packagingType) {
            return yieldData[i][2];
        }
    }
    
    console.error(`No yield data found for ${product} - ${packagingType}`);
    return 0;
}

function testCreatePlantingPlan() {
    const testDate = '2024-12-18';
    
    console.log('=== Starting Planting Plan Test ===');
    console.log('Testing date:', testDate);
    
    try {
        // Create logger with debug mode on
        const logger = new Logger(true, 'VERBOSE');
        logger.info('Starting planting plan test');

        // Process orders first to get the filtered orders
        const result = processOrders(testDate, logger);
        
        if (result.orderCount === 0) {
            console.log('No orders found for test date');
            return;
        }

        // Create planting plan
        const pdfBlob = createPlantingPlan(result.orders, testDate);
        console.log('Planting plan created successfully');
        
        // Email the plan (optional for testing)
        MailApp.sendEmail({
            to: 'forgetmenotfarm90@gmail.com',
            subject: `Test Planting Plan for ${testDate}`,
            body: 'This is a test planting plan.',
            attachments: [pdfBlob]
        });
        console.log('Test planting plan emailed');

    } catch (error) {
        console.error('Test failed:', error);
        console.error('Stack trace:', error.stack);
    }
    
    console.log('=== Test Complete ===');
}

function showPlantingDateSelector() {
    const template = HtmlService.createTemplateFromFile('DateSelector');
    template.action = 'plantingPlan';
    const html = template.evaluate()
        .setWidth(300)
        .setHeight(200);
    SpreadsheetApp.getUi().showModalDialog(html, 'Select Date for Planting Plan');
}

function createPlantingPlanFromDate(dateString) {
    const logger = new Logger(true, 'INFO');
    
    try {
        logger.info('Starting planting plan generation');
        logger.verbose(`Input date: ${dateString}`);
        
        // Process orders
        const result = processOrders(dateString, logger);
        
        if (result.orderCount === 0) {
            SpreadsheetApp.getUi().alert('No orders found for selected date');
            return;
        }

        // Create planting plan
        const pdfBlob = createPlantingPlan(result.orders, dateString);
        
        // Email the plan
        MailApp.sendEmail({
            to: 'forgetmenotfarm90@gmail.com',
            subject: `Planting Plan for ${dateString}`,
            body: `Please find attached the planting plan for ${dateString}.\n\nTotal Products: ${Object.keys(result.orders).length}`,
            attachments: [pdfBlob]
        });
        
        SpreadsheetApp.getUi().alert('Planting plan created and emailed successfully!');
        
    } catch (error) {
        logger.critical('Error creating planting plan', error);
        SpreadsheetApp.getUi().alert('Error: ' + error.message);
    }
}

function updateMetricsFromGrowData() {
    const metricsSheet = SpreadsheetApp.openById('14UbiYfjOonBZY-55RVPujLVKJsGBd43J6L-9q6ZiSOU')
        .getSheetByName('Yield & Seed Metrics');
    const growsSheet = SpreadsheetApp.openById('1PYPh1LfkAUggO2fZYytuDFH5qDzXcDkHIrKO7_qdxx4')
        .getSheetByName('Grows');

    const growData = growsSheet.getDataRange().getValues();
    const headers = growData[0];
    
    // Get column indices
    const varietyIndex = headers.indexOf('VARIETY');
    const seedLotIndex = headers.indexOf('SEED LOT');
    const seedDensityIndex = headers.indexOf('SEED DENSITY');
    const harvestYieldIndex = headers.indexOf('HARVEST YIELD');
    const statusIndex = headers.indexOf('Status');
    const harvestDateIndex = headers.indexOf('HARVEST DATE');

    // Process grow data
    const varietyStats = {};

    // Skip header row
    for (let i = 1; i < growData.length; i++) {
        const row = growData[i];
        const variety = row[varietyIndex];
        const seedLot = row[seedLotIndex];
        const seedDensity = row[seedDensityIndex];
        const harvestYield = row[harvestYieldIndex];
        const status = row[statusIndex];
        const harvestDate = row[harvestDateIndex] ? new Date(row[harvestDateIndex]) : null;
        
        if (status === 'Harvested' && harvestYield > 0) {
            if (!varietyStats[variety]) {
                varietyStats[variety] = {
                    overall: {
                        yields: [],
                        seedRates: []
                    },
                    byLot: {},
                    seasonal: {
                        spring: { yields: [], seedRates: [] },
                        summer: { yields: [], seedRates: [] },
                        fall: { yields: [], seedRates: [] },
                        winter: { yields: [], seedRates: [] }
                    }
                };
            }

            // Add to overall stats
            varietyStats[variety].overall.yields.push(harvestYield);
            varietyStats[variety].overall.seedRates.push(seedDensity);

            // Add to seed lot specific stats
            if (!varietyStats[variety].byLot[seedLot]) {
                varietyStats[variety].byLot[seedLot] = {
                    yields: [],
                    seedRates: []
                };
            }
            varietyStats[variety].byLot[seedLot].yields.push(harvestYield);
            varietyStats[variety].byLot[seedLot].seedRates.push(seedDensity);

            // Add to seasonal stats if harvest date exists
            if (harvestDate) {
                const month = harvestDate.getMonth();
                let season;
                if (month >= 2 && month <= 4) season = 'spring';
                else if (month >= 5 && month <= 7) season = 'summer';
                else if (month >= 8 && month <= 10) season = 'fall';
                else season = 'winter';

                varietyStats[variety].seasonal[season].yields.push(harvestYield);
                varietyStats[variety].seasonal[season].seedRates.push(seedDensity);
            }
        }
    }

    // Calculate averages and prepare data
    const metricsData = [];
    Object.entries(varietyStats).forEach(([variety, stats]) => {
        // Calculate overall averages
        const overallAvgYield = calculateAverage(stats.overall.yields);
        const overallAvgSeedRate = calculateAverage(stats.overall.seedRates);

        // Get current seed lot (most recent)
        const currentLot = Object.keys(stats.byLot).pop();
        const currentLotStats = stats.byLot[currentLot];
        const lotAvgYield = calculateAverage(currentLotStats.yields);
        const lotAvgSeedRate = calculateAverage(currentLotStats.seedRates);

        // Get seasonal averages
        const currentSeason = getCurrentSeason();
        const seasonalStats = stats.seasonal[currentSeason];
        const seasonalAvgYield = calculateAverage(seasonalStats.yields);
        const seasonalAvgSeedRate = calculateAverage(seasonalStats.seedRates);

        metricsData.push([
            variety,
            Math.round(overallAvgYield),
            Math.round(overallAvgSeedRate),
            currentLot,
            Math.round(lotAvgYield),
            Math.round(lotAvgSeedRate),
            currentSeason,
            Math.round(seasonalAvgYield),
            Math.round(seasonalAvgSeedRate),
            0,  // Weekly Usage (g)
            `Overall: ${stats.overall.yields.length} grows, Current Lot: ${currentLotStats.yields.length} grows, ${currentSeason}: ${seasonalStats.yields.length} grows`
        ]);
    });

    // Sort by variety name
    metricsData.sort((a, b) => a[0].localeCompare(b[0]));

    // Update sheet headers
    const headers = [
        ['Variety', 
         'Overall Yield/Tray (g)', 'Overall Seeds/Tray (g)',
         'Current Seed Lot', 'Lot Yield/Tray (g)', 'Lot Seeds/Tray (g)',
         'Current Season', 'Seasonal Yield/Tray (g)', 'Seasonal Seeds/Tray (g)',
         'Weekly Usage (g)', 'Notes'
        ]
    ];

    // Clear and update sheet
    metricsSheet.clear();
    metricsSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    if (metricsData.length > 0) {
        metricsSheet.getRange(2, 1, metricsData.length, metricsData[0].length)
            .setValues(metricsData);
    }

    // Format numbers
    const numberColumns = [2,3,5,6,8,9,10];  // columns with numeric data
    numberColumns.forEach(col => {
        metricsSheet.getRange(2, col, metricsData.length, 1).setNumberFormat('#,##0.0');
    });
}

function calculateAverage(numbers) {
    return numbers.length > 0 ? numbers.reduce((a, b) => a + b, 0) / numbers.length : 0;
}

function getCurrentSeason() {
    const month = new Date().getMonth();
    if (month >= 2 && month <= 4) return 'spring';
    if (month >= 5 && month <= 7) return 'summer';
    if (month >= 8 && month <= 10) return 'fall';
    return 'winter';
} 