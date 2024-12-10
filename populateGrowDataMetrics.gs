class MetricsGenerator {
    constructor(config = {}) {
        // Default configuration with overrides from passed config
        this.config = {
    DEBUG: true,
    LOG_LEVEL: 'VERBOSE',
    PLANTING_CYCLE: {
                PLANTING_DAY: 5,
                CYCLE_LENGTH: 7,
                CYCLE_OFFSET: 1,
                CYCLE_START_HOUR: 0,
                CYCLE_END_HOUR: 23,
            },
            ...config
        };

        // Initialize logger once for the instance
        this.logger = new Logger(this.config.DEBUG, this.config.LOG_LEVEL);
    }

    async generate() {
        try {
            this.logger.info('Starting grow data metrics population');
            
            const growData = await this.getGrowData();
            const varietyStats = await this.processGrowData(growData);
            const metrics = await this.generateMetricsData(varietyStats, growData);
            await this.updateMetricsSheet(metrics);

            return {
                success: true,
                logs: this.logger.logMessages
            };
        } catch (error) {
            this.logger.critical('Failed to populate metrics', error);
            return {
                success: false,
                error: error.message,
                logs: this.logger.logMessages
            };
        }
    }

    async getGrowData() {
        const growsSheet = SpreadsheetApp.openById('1PYPh1LfkAUggO2fZYytuDFH5qDzXcDkHIrKO7_qdxx4')
            .getSheetByName('Grows');

        if (!growsSheet) {
            throw new Error('Grows sheet not found');
        }

        const growData = growsSheet.getDataRange().getValues();
        const indices = this.findGrowSheetColumns(growData[1]);
        
        return { data: growData, indices };
    }

    // ... other methods would be class methods with access to this.logger
} 