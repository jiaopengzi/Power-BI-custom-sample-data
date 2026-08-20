/**
 * FilePath    : frontend/src/i18n/en-us.ts
 * Author      : jiaopengzi
 * Blog        : https://jiaopengzi.com
 * Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
 * Description : 英文文案.
 */

import type { Messages } from "./types"

// English messages.
const messages: Messages = {
    title: "Power BI Sample Data Generator",
    subtitle: "Power BI Sample Data Generator",
    form: {
        startDate: "Start Date",
        endDate: "End Date",
        productCount: "Products",
        storeCount: "Stores",
        inventoryCycle: "Inventory Cycle",
        outputDir: "Output Directory",
        outputDirPlaceholder: "Please choose an output directory first",
        localeLabel: "Language",
    },
    buttons: {
        generate: "Generate Sample Data",
        incremental: "Incremental Update",
        chooseDir: "Choose Directory",
        openDir: "Open Directory",
        docs: "Documentation",
    },
    stages: {
        stageStart: "Preparing...",
        stageDimensions: "Generating dimensions (region/province/city/district)...",
        stageProducts: "Generating products...",
        stageStores: "Generating stores...",
        stageCustomers: "Generating customers...",
        stageOrders: "Generating inventory/orders...",
        stageDone: "Done!",
    },
    result: {
        title: "Result",
        products: "Products",
        stores: "Stores",
        customers: "Customers",
        inventory: "Inventory",
        orders: "Orders",
        orderItem: "Order Items",
        rows: "rows",
    },
    msg: {
        chooseDirFirst: "Please choose an output directory first",
        invalidRange: "End date must be after start date, window at least 60 days",
        generating: "Generating, please wait...",
        generateSuccess: "Sample data generated",
        incrementalSuccess: "Incremental update completed",
        noBaseData: 'No base data in directory. Click "Generate Sample Data" first',
        dateConflict: "Incremental range conflicts with existing data: {range}. Adjust dates",
        failed: "Operation failed: {msg}",
        productRange: "Products must be between 1 and 2000",
        storeRange: "Stores must be between 1 and 400",
        inventoryRange: "Inventory cycle must be between 5 and 20",
        overwriteTitle: "Overwrite existing data?",
        overwriteContent: "The output directory already contains sample data. Generating will overwrite it. Continue?",
        overwriteConfirm: "Overwrite",
        overwriteCancel: "Cancel",
    },
}

export default messages
