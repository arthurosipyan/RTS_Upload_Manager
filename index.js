function upload() {
    document.getElementById("copyTextButton").style.display = "block";
    let files = document.getElementById('fileUpload').files;
    
    if (files.length==0) {
        alert("Please choose any file...");
        return;
    }

    let filename = files[0].name;
    let extension = filename.substring(filename.lastIndexOf(".")).toUpperCase();
    
    if (extension == '.XLS' || extension == '.XLSX') {
        processFile(files[0]);
    }
    else {
        alert("Please select a valid excel file.");
    }
}

function copyText() {
    /* Get the text field */
    let copiedText = document.getElementById("sheetResults");
  
    /* Select the text field */
    copiedText.select();
    copiedText.setSelectionRange(0, 99999); /* For mobile devices */
    
    /* Copy the text inside the text field */
    navigator.clipboard.writeText(copiedText.value);
  
    /* Alert the copied text */
    alert("Copied the text: " + copiedText.value);
}

function getTotalRows(invoices) {
    return invoices.length;
}

function getTotalAmount(sheet) {
    let x = 0;
    let total = 0;
    while (x < sheet.length) {
        total += sheet[x]['InvAmt'];
        x++;
    }
    return total;
}

function getInvoices(sheet) {
    let x = 0;
    let invoiceArr = [];
    while (x < sheet.length) {
        invoiceArr.push('"' + sheet[x]['Invoice#'] + '"');
        x++;
    }
    invoiceArr = invoiceArr.join(' ');
    return invoiceArr;
}

// Method to read and process excel file
function processFile(file) {
    try {
        let reader = new FileReader();
        reader.readAsBinaryString(file);
        reader.onload = function(e) {
            let data = e.target.result;
            let workbook = XLSX.read(data, {
                type : 'binary'
            });
            let lastSheet = workbook.SheetNames[workbook.SheetNames.length-1];
            
            // store sheets in JSON
            let result = {};
            workbook.SheetNames.forEach(function(sheetName) {
                let roa = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
                if (roa.length > 0) {
                    result[sheetName] = roa;
                }
            });

            // get list of invoices
            let invoiceList = getInvoices(result[lastSheet]);

            // get total invoice count
            let totalInvoices = getTotalRows(result[lastSheet]);

            // get total upload amount
            let totalAmount = (getTotalAmount(result[lastSheet])).toLocaleString('en-US', {
                style: 'currency',
                currency: 'USD',
            });;
            
            // display results
            var sheetStats = document.getElementById("totalInvoices");
            sheetStats.textContent = "TOTAL INVOICES: " + totalInvoices;

            var sheetStats = document.getElementById("totalAmount");
            sheetStats.textContent = "TOTAL AMOUNT: " + totalAmount;

            var resultEle=document.getElementById("sheetResults");
            resultEle.textContent = invoiceList;
            resultEle.style.display='block';
        }
    
    }catch(e) {
        console.error(e);
    }
}