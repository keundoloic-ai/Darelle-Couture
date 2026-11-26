
// ======================================
// DARELLE COUTURE - BUSINESS ANALYTICS
// ======================================

// JavaScript connection test

console.log("Darelle Couture Analytics loaded successfully!");

console.log("JavaScript is working correctly!");

// SheetJS connection test

if (typeof XLSX !== "undefined") {
    console.log("SheetJS loaded successfully!");
    console.log("SheetJS version:", XLSX.version);
} else {
    console.error("SheetJS failed to load!");
}


// ======================================
// CHART.JS CONNECTION TEST
// ======================================

if (typeof Chart !== "undefined") {

    console.log("Chart.js loaded successfully!");

    console.log("Chart.js version:", Chart.version);

} else {

    console.error("Chart.js failed to load!");

}

// ======================================
// REGISTER CHART.JS DATALABELS PLUGIN
// ======================================

if (typeof ChartDataLabels !== "undefined" &&
    typeof Chart !== "undefined") {

    Chart.register(ChartDataLabels);

    console.log("DataLabels loaded successfully!");

} else {

    console.warn("DataLabels CDN unavailable; using built-in fallback.");
    if (typeof Chart !== "undefined") {
        Chart.register({
            id: "fallbackValueLabels",
            afterDatasetsDraw(chart) {
                const options = chart.options.plugins.datalabels;
                if (!options) return;
                const ctx = chart.ctx;
                ctx.save();
                ctx.font = "bold 14px Arial";
                ctx.fillStyle = "#24201d";
                ctx.textAlign = "center";
                ctx.textBaseline = "bottom";
                chart.data.datasets.forEach((dataset, index) => {
                    if (!chart.isDatasetVisible(index)) return;
                    chart.getDatasetMeta(index).data.forEach((bar, i) => {
                        const value = dataset.data[i];
                        const label = typeof options.formatter === "function"
                            ? options.formatter(value) : value;
                        ctx.fillText(String(label), bar.x, bar.y - 5);
                    });
                });
                ctx.restore();
            }
        });
    }

}

// ======================================
// LOAD EXCEL FILE
// ======================================

async function loadExcelData() {

    try {

        const response = await fetch(
            "Data/Darelle_IS_Analys.xlsx"
        );

        if (!response.ok) {
            throw new Error(
                "HTTP Error: " + response.status
            );
        }

        const data = await response.arrayBuffer();

        const workbook = XLSX.read(data, {
            type: "array"
        });

        console.log("Excel file loaded successfully!");

        console.log("Available worksheets:",
            workbook.SheetNames
        );


// Select the raw data worksheet

const worksheet = workbook.Sheets[
    "Darelle_IS_Analysis xlsx"
];

// Check if the worksheet exists

if (!worksheet) {
    throw new Error("Worksheet not found!");
}

// Convert Excel data into JavaScript objects

const records = XLSX.utils.sheet_to_json(worksheet);

// Display the extracted data

console.log("Excel data extracted successfully!");

console.log("Total records:", records.length);

// Display only the fields needed for analytics

console.table(records.map(record => ({
    service_name: record.service_name,
    base_price: record.base_price,
    order_id: record.order_id
})));


// ======================================
// CREATE DYNAMIC PIVOT TABLE
// ======================================

const pivotTable = {};

// Group records by service

records.forEach(record => {

    const service = record.service_name;
    const price = Number(record.base_price);

    // Create a new service category

    if (!pivotTable[service]) {

        pivotTable[service] = {
            totalOrders: 0,
            serviceValue: 0
        };

    }

    // Count orders

    pivotTable[service].totalOrders += 1;

    // Calculate service value

    pivotTable[service].serviceValue += price;

});

// Display PivotTable results

console.log("PivotTable created successfully!");

console.table(pivotTable);


// ======================================
// DISPLAY PIVOT TABLE ON WEBSITE
// ======================================

const tableBody = document.getElementById(
    "pivot-table-body"
);

// Clear existing table data

tableBody.innerHTML = "";

// Canadian dollar formatter

const currency = new Intl.NumberFormat("en-CA", {
    style: "currency",
    currency: "CAD"
});

// Insert each service into the HTML table

Object.entries(pivotTable).forEach(([service, values]) => {

    const row = document.createElement("tr");

    const serviceName = service
        .replaceAll("_", " ")
        .replace(/\b\w/g, letter => letter.toUpperCase());

    row.insertCell().textContent = serviceName;

    row.insertCell().textContent = values.totalOrders;

    row.insertCell().textContent =
        currency.format(values.serviceValue);

    tableBody.appendChild(row);

});

console.log("PivotTable displayed on website successfully!");


// ======================================
// CALCULATE GRAND TOTAL
// ======================================

let totalOrders = 0;
let totalServiceValue = 0;

// Calculate totals

Object.values(pivotTable).forEach(values => {

    totalOrders += values.totalOrders;

    totalServiceValue += values.serviceValue;

});

// Create the Grand Total row

const totalRow = document.createElement("tr");

totalRow.classList.add("grand-total");

totalRow.insertCell().textContent = "Grand Total";

totalRow.insertCell().textContent = totalOrders;

totalRow.insertCell().textContent =
    currency.format(totalServiceValue);

// Add the total row to the table

tableBody.appendChild(totalRow);

// Display totals in the console

console.log("Total Orders:", totalOrders);

console.log("Total Service Value:", totalServiceValue);

console.log("Grand Total calculated successfully!");

// ======================================
// UPDATE KPI CARDS
// ======================================

// Display total orders

document.getElementById("kpi-orders").textContent =
    totalOrders;

// Display total service value

document.getElementById("kpi-revenue").textContent =
    currency.format(totalServiceValue);

// Display number of service categories

document.getElementById("kpi-categories").textContent =
    Object.keys(pivotTable).length;

// Confirmation message

console.log("KPI cards updated successfully!");


// ======================================
// CREATE ORDERS BY SERVICE CHART
// ======================================

// Extract service categories

const chartLabels = Object.keys(pivotTable).map(service =>
    service
        .replaceAll("_", " ")
        .replace(/\b\w/g, letter => letter.toUpperCase())
);

// Extract total orders

const chartOrders = Object.values(pivotTable).map(
    values => values.totalOrders
);

// Select the HTML canvas

const ordersCanvas = document.getElementById("ordersChart");

// Create the chart


const ordersChart = new Chart(ordersCanvas, {

    type: "bar",

    data: {
        labels: chartLabels,

        datasets: [{
            label: "Total Orders",
            data: chartOrders,
            backgroundColor: "#a67c52",
            borderColor: "#a67c52",
            borderWidth: 1,
            borderRadius: 5
        }]
    },

    options: {

        responsive: true,
        maintainAspectRatio: false,

        scales: {
            y: {
                beginAtZero: true,
                grace: "20%",
                ticks: {
                    precision: 0
                }
            }
        },

        plugins: {

            legend: {
                display: false
            },

            datalabels: {
                anchor: "end",
                align: "top",
                color: "#24201d",

                font: {
                    size: 14,
                    weight: "bold"
                },

                formatter: function(value) {
                    return value;
                }
            },

            tooltip: {
                backgroundColor: "#24201d",
                padding: 12,
                cornerRadius: 8
            }

        }

    }

});

console.log("Orders chart created successfully!");

// ======================================
// CREATE SERVICE VALUE CHART
// ======================================

// Extract service values

const chartRevenue = Object.values(pivotTable).map(
    values => values.serviceValue
);

// Select the HTML canvas

const revenueCanvas = document.getElementById(
    "revenueChart"
);

// Create the revenue chart


// ======================================
// REVENUE CHART CONFIGURATION
// ======================================

const revenueChart = new Chart(revenueCanvas, {

    type: "bar",

    data: {

        labels: chartLabels,

        datasets: [{

            label: "Service Value (CAD)",

            data: chartRevenue,

            backgroundColor: "#24201d",

            borderColor: "#24201d",

            borderWidth: 1,

            borderRadius: 5

        }]

    },

    options: {

        responsive: true,

        maintainAspectRatio: false,

        scales: {

            y: {

                beginAtZero: true,

                grace: "20%",

                ticks: {

                    callback: function(value) {

                        return currency.format(value);

                    }

                }

            }

        },

        plugins: {

            legend: {
                display: false
            },

            datalabels: {

                anchor: "end",

                align: "top",

                color: "#24201d",

                font: {
                    size: 14,
                    weight: "bold"
                },

                formatter: function(value) {

                    return currency.format(value);

                }

            },

            tooltip: {

                backgroundColor: "#24201d",

                padding: 12,

                cornerRadius: 8,

                callbacks: {

                    label: function(context) {

                        return "Service Value: " +
                            currency.format(context.parsed.y);

                    }

                }

            }

        }

    }

});

console.log("Revenue chart created successfully!");

    } catch (error) {

        console.error("Error loading Excel file:", error);

    }
}

loadExcelData();

// ======================================
// BUILD UPLOADED PIVOT TABLE
// ======================================

function buildUploadedPivotTable(data) {

    const newPivotTable = Object.create(null);

    data.forEach(function(row) {

        const serviceName = row.service_name.trim();

        const servicePrice = Number(row.base_price);

        // Create service category if it does not exist

        if (!newPivotTable[serviceName]) {

            newPivotTable[serviceName] = {

                totalOrders: 0,

                serviceValue: 0

            };

        }

        // Count orders

        newPivotTable[serviceName].totalOrders += 1;

        // Calculate service value

        newPivotTable[serviceName].serviceValue += servicePrice;

    });

    console.log("Uploaded PivotTable created successfully!");

    console.table(newPivotTable);

    return newPivotTable;

}


// ======================================
// UPDATE DASHBOARD FROM UPLOADED DATA
// ======================================

function updateUploadedDashboard(newPivotTable) {

    const tableBody = document.getElementById(
        "pivot-table-body"
    );

    const currency = new Intl.NumberFormat("en-CA", {
        style: "currency",
        currency: "CAD"
    });

    // Clear previous table data

    tableBody.innerHTML = "";

    let totalOrders = 0;
    let totalServiceValue = 0;

    // Display new PivotTable data

    Object.entries(newPivotTable).forEach(
        ([service, values]) => {

            const serviceName = service
                .replace(/_/g, " ")
                .replace(/\b\w/g, letter => letter.toUpperCase());

            const row = tableBody.insertRow();

            row.insertCell().textContent = serviceName;

            row.insertCell().textContent = values.totalOrders;

            row.insertCell().textContent =
                currency.format(values.serviceValue);

            totalOrders += values.totalOrders;

            totalServiceValue += values.serviceValue;

        }
    );

    // Create Grand Total row

    const totalRow = tableBody.insertRow();

    totalRow.classList.add("grand-total");

    totalRow.insertCell().textContent = "Grand Total";

    totalRow.insertCell().textContent = totalOrders;

    totalRow.insertCell().textContent =
        currency.format(totalServiceValue);

    // Update KPI cards

    document.getElementById("kpi-orders").textContent =
        totalOrders;

    document.getElementById("kpi-revenue").textContent =
        currency.format(totalServiceValue);

    document.getElementById("kpi-categories").textContent =
        Object.keys(newPivotTable).length;

    // Display results in console

    console.log("Dynamic PivotTable updated successfully!");

    console.log("Dynamic KPI cards updated successfully!");

    console.log("Updated Total Orders:", totalOrders);

    console.log("Updated Service Value:", totalServiceValue);

}

// ======================================
// UPDATE DYNAMIC PIVOT CHARTS
// ======================================

function updateUploadedCharts(newPivotTable) {

    // Extract service categories

    const labels = Object.keys(newPivotTable).map(
        service => service
            .replace(/_/g, " ")
            .replace(/\b\w/g, letter => letter.toUpperCase())
    );

    // Extract orders

    const ordersData = Object.values(newPivotTable).map(
        service => service.totalOrders
    );

    // Extract service values

    const revenueData = Object.values(newPivotTable).map(
        service => service.serviceValue
    );

    // Retrieve existing charts

    const ordersChart = Chart.getChart("ordersChart");

    const revenueChart = Chart.getChart("revenueChart");

    if (!ordersChart || !revenueChart) {

        throw new Error("Existing charts could not be found.");

    }

    // Update Orders Chart

    ordersChart.data.labels = labels;

    ordersChart.data.datasets[0].data = ordersData;

    ordersChart.update();

    // Update Revenue Chart

    revenueChart.data.labels = labels;

    revenueChart.data.datasets[0].data = revenueData;

    revenueChart.update();

    console.log("Dynamic Orders Chart updated successfully!");

    console.log("Dynamic Revenue Chart updated successfully!");

}

// ======================================
// DYNAMIC EXCEL FILE IMPORT
// ======================================

const excelFileInput = document.getElementById("excelFile");

const fileNameDisplay = document.getElementById("fileName");

const uploadStatus = document.getElementById("uploadStatus");

excelFileInput.addEventListener("change", async function(event) {

    const file = event.target.files[0];

    if (!file) {
        return;
    }

    fileNameDisplay.textContent = file.name;

    uploadStatus.textContent = "Validating Excel file...";
    uploadStatus.style.color = "#a67c52";

    try {


// ======================================
// FILE SIZE VALIDATION
// ======================================

const maxFileSize = 5 * 1024 * 1024;

if (file.size > maxFileSize) {

    throw new Error(
        "File is too large. Maximum size is 5 MB."
    );

}

        // Validate file extension

        if (!file.name.toLowerCase().endsWith(".xlsx")) {

            throw new Error("Only .xlsx files are supported.");

        }

        // Read Excel file

        const fileData = await file.arrayBuffer();

        const uploadedWorkbook = XLSX.read(fileData, {
            type: "array"
        });

        console.log("Uploaded Excel file loaded successfully!");

        // Required columns

        const requiredColumns = [
            "service_name",
            "base_price",
            "order_id"
        ];

        let uploadedData = [];
        let validSheetName = null;

        // Search for the correct worksheet

        for (const sheetName of uploadedWorkbook.SheetNames) {

            const worksheet = uploadedWorkbook.Sheets[sheetName];

            const rows = XLSX.utils.sheet_to_json(worksheet, {
                defval: null
            });

            if (rows.length === 0) {
                continue;
            }

            const columns = Object.keys(rows[0]);

            const isValid = requiredColumns.every(
                column => columns.includes(column)
            );

            if (isValid) {

                uploadedData = rows;
                validSheetName = sheetName;

                break;

            }

        }

        // Validate worksheet

        if (!validSheetName) {

            throw new Error(
                "Required columns are missing from the Excel file."
            );

        }

// ======================================
// RECORD LIMIT VALIDATION
// ======================================

const maxRecords = 10000;

if (uploadedData.length > maxRecords) {

    throw new Error(
        "Maximum of 10,000 records allowed."
    );

}


        // Validate individual records

        uploadedData.forEach(function(row, index) {

            if (
                typeof row.service_name !== "string" ||
                row.service_name.trim() === "" ||
                typeof row.base_price !== "number" ||
                !Number.isFinite(row.base_price) ||
                row.base_price < 0 ||
                row.order_id === null ||
                String(row.order_id).trim() === ""
            ) {

                throw new Error(
                    "Invalid data in record " + (index + 1)
                );

            }

        });

        // Display validation results

        console.log("Valid worksheet:", validSheetName);

        console.log("Excel validation successful!");

// ======================================
// RECALCULATE PIVOT TABLE
// ======================================

const newPivotTable = buildUploadedPivotTable(uploadedData);

// UPDATE DASHBOARD

updateUploadedDashboard(newPivotTable);

// UPDATE DASHBOARD

updateUploadedDashboard(newPivotTable);

console.log("Dynamic PivotTable recalculated successfully!");


        console.log("Total imported records:", uploadedData.length);

        console.table(uploadedData.map(row => ({
            service_name: row.service_name,
            base_price: row.base_price,
            order_id: row.order_id
        })));

        uploadStatus.textContent =
            "Excel validated successfully! " +
            uploadedData.length + " records imported.";

        uploadStatus.style.color = "green";

    } catch (error) {

        console.error("Excel import error:", error.message);

        uploadStatus.textContent = error.message;

        uploadStatus.style.color = "red";

    } finally {

        excelFileInput.value = "";

    }

});

