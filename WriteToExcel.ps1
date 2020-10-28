try {
    Add-Type -Path "$PSScriptRoot\ClosedXML.dll"
    Add-Type -Path "$PSScriptRoot\DocumentFormat.OpenXml.dll"
    
    #Create the excel workbook
    $workBook = new-object ClosedXML.Excel.XLWorkbook
    #Add a new sheet and some data
    $workSheet = $workBook.Worksheets.Add("Test")   
    $worksheet.Cell("A1").Value = "No Data";
    #Save the workbook
    $workBook.SaveAs("$PSScriptRoot\test.xlsx")
}

catch [Exception] {
    Write-Host $_.Exception.Message
    exit 1
}