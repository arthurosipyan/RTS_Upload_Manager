// ClientNo	Invoice#	DebtorNo	DebtorName	Pono	InvDate	 InvAmt 

// Method to upload a valid excel file
function upload() {
    var files = document.getElementById('fileUpload').files;
    if(files.length==0){
        alert("Please choose any file...");
        return;
    }
    var filename = files[0].name;
    var extension = filename.substring(filename.lastIndexOf(".")).toUpperCase();
    if (extension == '.XLS' || extension == '.XLSX') {
        excelFileToJSON(files[0]);
    }else{
        alert("Please select a valid excel file.");
    }
  }

// Method to read excel file and convert it into JSON 
function excelFileToJSON(file){
    try {
        var reader = new FileReader();
        reader.readAsBinaryString(file);
        reader.onload = function(e) {
            var data = e.target.result;
            var workbook = XLSX.read(data, {
                type : 'binary'
            });
            // let workbookSheets = workbook.SheetNames;
            let lastSheet = workbook.SheetNames[workbook.SheetNames.length-1];
            
            // store sheets in result object
            var result = {};
            workbook.SheetNames.forEach(function(sheetName) {
                var roa = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
                if (roa.length > 0) {
                    result[sheetName] = roa;
                }
              });

            // var result = XLSX.utils.sheet_to_row_object_array(workbookSheets[workbookSheets.length-1]);
            
            // displaying the json result
            var resultEle=document.getElementById("json-result");
            resultEle.value=JSON.stringify(result[lastSheet], null, 4);
            resultEle.style.display='block';
            }
      }catch(e){
          console.error(e);
      }
}