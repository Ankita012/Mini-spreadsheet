// Define the spreadsheet dimensions
const rows = 100;
const cols = 100;

// Initialize the data object to store cell values
let data = {};

// Function to generate column labels (A to Z, then AA, AB, AC, etc.)
function getColumnLabel(col) {
    let label = "";
    while (col > 0) {
        const remainder = col % 26 || 26;
        label = String.fromCharCode(64 + remainder) + label;
        col = Math.floor((col - 1) / 26);
    }
    return label;
}

// Function to generate the spreadsheet
function createSpreadsheet() {
    const table = document.getElementById("spreadsheet");
    table.innerHTML = "";
    const thead = document.createElement("thead");
    const tr = document.createElement("tr");
    tr.appendChild(document.createElement("th"));
    for (let j = 0; j < cols; j++) {
        const th = document.createElement("th");
        th.textContent = getColumnLabel(j + 1);
        tr.appendChild(th);
    }
    thead.appendChild(tr);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");
    for (let i = 0; i < rows; i++) {
        const tr = document.createElement("tr");
        const th = document.createElement("th");
        th.textContent = i + 1;
        tr.appendChild(th);
        for (let j = 0; j < cols; j++) {
            const td = document.createElement("td");
            const input = document.createElement("input");
            input.type = "text";
            input.dataset.row = i;
            input.dataset.col = j;
            input.addEventListener("input", updateCell);
            td.appendChild(input);
            tr.appendChild(td);
        }
        tbody.appendChild(tr);
    }
    table.appendChild(tbody);
}

// Function to update the cell value
function updateCell(event) {
    const input = event.target;
    const row = parseInt(input.dataset.row);
    const col = parseInt(input.dataset.col);
    const value = input.value.trim();
    if (!data[row]) {
        data[row] = {};
    }
    data[row][col] = value;
}

// Function to evaluate the formula on Enter key press
function handleEnterKeyPress(event) {
    const input = event.target;
    if (event.key === "Enter" && input.value.startsWith("=")) {
        const row = parseInt(input.dataset.row);
        const col = parseInt(input.dataset.col);
        try {
            const result = evaluateFormula(input.value.substring(1));
            data[row][col] = result;
            input.value = result;
        } catch (error) {
            input.value = "Error";
        }
    }else if (event.key === "Enter" && !input.value.startsWith("=")) {
        alert('please put = before formula');
    }
}

// Function to evaluate a formula
function evaluateFormula(formula) {
        const tokens = formula.match(/[A-Z]+\d+|\d+|\+|\-|\*|\/|sum|average|max|\:[A-Z]+\d+/gi);
        let result = 0;
        let currentOperator = null;

        for (let i = 0; i < tokens.length; i++) {
            const token = tokens[i];
            if (/^[A-Z]+\d+$/.test(token)) {
                const [col, row] = parseCell(token);
                const value = parseFloat(data[row][col]) || 0;
                result = applyOperator(result, value, currentOperator);
            } else if (/^\d+$/.test(token)) {
                const value = parseFloat(token);
                result = applyOperator(result, value, currentOperator);
            } else if(/^(sum|average|max|min)$/i.test(token))  {
                
                const rangeIndex = tokens.indexOf(token) + 1;
                const range = tokens.slice(rangeIndex).join(" ");

                if(/^sum$/i.test(token))
                    var ans = supportFunction(range, 'sum');
                else if(/^average$/i.test(token))
                   var ans = supportFunction(range, 'avg');
                else if(/^max$/i.test(token))
                    var ans = supportFunction(range, 'max');
                else if(/^min$/i.test(token))
                    var ans = supportFunction(range, 'min');

                result = applyOperator(result, ans, currentOperator);
                i += range.split(" ").length;
            }
            else {
                currentOperator = token;
            }
        }

    return result;
}
// Function to apply the operator to operands
function applyOperator(operand1, operand2, operator) {
    switch (operator) {
        case "+":
            return operand1 + operand2;
        case "-":
            return operand1 - operand2;
        case "*":
            return operand1 * operand2;
        case "/":
            return operand1 / operand2;
        default:
            return operand2;
    }
}
// support function for some basic operation
function supportFunction(range, type){
    const [startCell, endCell] = range.split(":");
    const [startCol, startRow] = parseCell(startCell);
    const [endCol, endRow] = parseCell(endCell);
    
    switch(type) {
    case 'sum' :
        return evaluateSum(startRow,endRow, startCol, endCol);
    case 'avg' :
        return evaluateAverage(startRow,endRow, startCol, endCol);
    case 'max' :
        return evaluateMax(startRow,endRow, startCol, endCol);
    case 'min' :
        return evaluateMin(startRow,endRow, startCol, endCol);
    }
}
// Function to evaluate the SUM function
function evaluateSum(startRow,endRow, startCol, endCol) {
    let sum = 0;
    for (let row = startRow; row <= endRow; row++) {
        for (let col = startCol; col <= endCol; col++) {
            sum += parseFloat(data[row][col]) || 0;
        }
    }
    return sum;
}
// Function to evaluate the AVERAGE function
function evaluateAverage(startRow,endRow, startCol, endCol) {
    let sum = 0;
    let count = 0;
    for (let row = startRow; row <= endRow; row++) {
        for (let col = startCol; col <= endCol; col++) {
            sum += parseFloat(data[row][col]) || 0;
            count++;
        }
    }
    const average = count > 0 ? sum / count : 0;
    return average;
}
// Function to evaluate the MAX function
function evaluateMax(startRow,endRow, startCol, endCol) {
    let max = parseFloat(data[startRow][startCol]) || 0;
    for (let row = startRow; row <= endRow; row++) {
        for (let col = startCol; col <= endCol; col++) {
            const value = parseFloat(data[row][col]) || 0;
            if (value > max) {
                max = value;
            }
        }
    }
    return max;
}
// Function to evaluate the MIN function
function evaluateMin(startRow,endRow, startCol, endCol) {
    let min = parseFloat(data[startRow][startCol]) || 0;
    for (let row = startRow; row <= endRow; row++) {
        for (let col = startCol; col <= endCol; col++) {
            const value = parseFloat(data[row][col]) || 0;
            if (value < min) {
                min = value;
            }
        }
    }
    return min;
}

// Function to parse cell reference
function parseCell(cell) {
    const match = cell.match(/([A-Z]+)(\d+)/);
    if (match) {
        const col = getColumnIndex(match[1]);
        const row = parseInt(match[2]) - 1;
        return [col, row];
    }
}

// Function to get column index from column label
function getColumnIndex(col) {
    let index = 0;
    for (let i = 0; i < col.length; i++) {
        index = index * 26 + col.charCodeAt(i) - 64;
    }
    return index - 1;
}

// Function to insert stored data into cells
function insertData() {
    for (const row in data) {
        for (const col in data[row]) {
            const cell = document.querySelector(`#spreadsheet input[data-row='${row}'][data-col='${col}']`);
            if (cell) {
                cell.value = data[row][col];
            }
        }
    }
}

// Refresh button event listener
document.getElementById("refreshButton").addEventListener("click", function() {
    createSpreadsheet();
    insertData();
});


// Add event listener for input changes
document.querySelectorAll("input").forEach(input => {
    input.addEventListener("input", updateCell);
});

// Add event listener for Enter key press
document.addEventListener("keyup", handleEnterKeyPress);

// Initialize the spreadsheet
createSpreadsheet();