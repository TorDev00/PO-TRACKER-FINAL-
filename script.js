// Handle file upload
document
  .getElementById("fileUpload")
  .addEventListener("change", handleFileUpload, false);

function handleFileUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = e.target.result;
    const workbook = XLSX.read(data, { type: "binary" });

    // Get the first sheet
    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    // Convert sheet to JSON
    const jsonData = XLSX.utils.sheet_to_json(sheet);

    console.log("Parsed Excel Data:", jsonData); // Log the data for debugging

    // Display PO Numbers and Vendors
    displayPOTable(jsonData);
  };
  reader.readAsBinaryString(file);
}

function displayPOTable(jsonData) {
  const poTableBody = document
    .getElementById("poTable")
    .getElementsByTagName("tbody")[0];
  poTableBody.innerHTML = ""; // Clear any existing rows
  const uniquePONumbers = new Set();

  // Separate POs with mismatched quantities and those without
  const mismatchedPOs = [];
  const matchedPOs = [];

  jsonData.forEach((row) => {
    const quantity = parseInt(row["QUANTITY"], 10) || 0;
    const received = parseInt(row["RECEIVED"], 10) || 0;

    if (quantity !== received) {
      mismatchedPOs.push(row); // Add to mismatched POs
    } else {
      matchedPOs.push(row); // Add to matched POs
    }
  });

  // Combine the arrays with mismatched POs first
  const sortedData = [...mismatchedPOs, ...matchedPOs];

  // Loop through the sorted data
  sortedData.forEach((row) => {
    if (!uniquePONumbers.has(row["PO NUMBER"])) {
      uniquePONumbers.add(row["PO NUMBER"]);

      // Create a row for the PO
      const poRow = document.createElement("tr");
      poRow.classList.add("po-row");

      // Check if quantity and received are different
      const quantity = parseInt(row["QUANTITY"], 10) || 0;
      const received = parseInt(row["RECEIVED"], 10) || 0;
      if (quantity !== received) {
        poRow.classList.add("highlight-po"); // Add highlight class to the PO row if mismatched
      }

      // PO Number cell
      const poCell = document.createElement("td");
      poCell.textContent = row["PO NUMBER"] || "N/A"; // Show 'N/A' if PO number is missing
      poRow.appendChild(poCell);

      // Vendor cell
      const vendorCell = document.createElement("td");
      vendorCell.textContent = row["VENDOR"] || "Unknown Vendor"; // Show 'Unknown Vendor' if vendor is missing
      poRow.appendChild(vendorCell);

      // Status cell
      const statusCell = document.createElement("td");
      statusCell.textContent = row["STATUS"] || "No Status"; // Show 'No Status' if status is missing
      poRow.appendChild(statusCell);

      // Dropdown arrow cell (shows more details on click)
      const dropdownCell = document.createElement("td");
      const dropdownBtn = document.createElement("button");
      dropdownBtn.classList.add("dropdown-btn");
      dropdownBtn.textContent = "▼"; // Arrow icon
      dropdownBtn.addEventListener("click", () =>
        toggleDetails(poRow, jsonData, row["PO NUMBER"])
      );
      dropdownCell.appendChild(dropdownBtn);
      poRow.appendChild(dropdownCell);

      // Add row to the table
      poTableBody.appendChild(poRow);

      // Add hidden details row with Product, Quantity, Received, and Left
      const detailsRow = document.createElement("tr");
      detailsRow.classList.add("details-row");
      const detailsCell = document.createElement("td");
      detailsCell.colSpan = 5; // Adjust colspan based on number of columns
      const detailsTable = createDetailsTable(
        jsonData.filter((item) => item["PO NUMBER"] === row["PO NUMBER"])
      );
      detailsCell.appendChild(detailsTable);
      detailsRow.appendChild(detailsCell);
      poTableBody.appendChild(detailsRow);
    }
  });
}

// Create the details table for a specific PO number
function createDetailsTable(items) {
  const detailsTable = document.createElement("table");
  detailsTable.classList.add("details-table");
  detailsTable.style.width = "100%";
  detailsTable.style.borderCollapse = "collapse";

  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");

  const productHeader = document.createElement("th");
  productHeader.textContent = "Product";
  headerRow.appendChild(productHeader);

  const quantityHeader = document.createElement("th");
  quantityHeader.textContent = "Quantity";
  headerRow.appendChild(quantityHeader);

  const receivedHeader = document.createElement("th");
  receivedHeader.textContent = "Received";
  headerRow.appendChild(receivedHeader);

  const leftHeader = document.createElement("th");
  leftHeader.textContent = "Left"; // Header for "Left" column
  headerRow.appendChild(leftHeader);

  thead.appendChild(headerRow);
  detailsTable.appendChild(thead);

  const tbody = document.createElement("tbody");
  items.forEach((item) => {
    const row = document.createElement("tr");
    const productCell = document.createElement("td");
    productCell.textContent = item["PRODUCT"] || "N/A"; // Show 'N/A' if product is missing
    row.appendChild(productCell);

    const quantityCell = document.createElement("td");
    quantityCell.textContent = item["QUANTITY"] || "0"; // Show '0' if quantity is missing
    row.appendChild(quantityCell);

    const receivedCell = document.createElement("td");
    receivedCell.textContent = item["RECEIVED"] || "0"; // Show '0' if received is missing
    row.appendChild(receivedCell);

    // Calculate the remaining (LEFT) pieces
    const quantity = parseInt(item["QUANTITY"], 10) || 0;
    const received = parseInt(item["RECEIVED"], 10) || 0;
    const left = quantity - received;

    // LEFT cell
    const leftCell = document.createElement("td");
    leftCell.textContent = left >= 0 ? left : "0"; // Show remaining pieces
    row.appendChild(leftCell);

    // Highlight mismatched received quantities in details table
    if (quantity !== received) {
      receivedCell.classList.add("highlight"); // Add highlight class if mismatched
    }

    tbody.appendChild(row);
  });

  detailsTable.appendChild(tbody);
  return detailsTable;
}

// Toggle the visibility of the details row
function toggleDetails(poRow, jsonData, poNumber) {
  const detailsRow = poRow.nextElementSibling; // The row immediately after the current PO row
  if (detailsRow.style.display === "table-row") {
    detailsRow.style.display = "none";
  } else {
    detailsRow.style.display = "table-row";
  }
}
