Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.Office.Interop.Word
Add-Type -AssemblyName Microsoft.Office.Interop.Excel
Add-Type -AssemblyName Microsoft.Office.Interop.Outlook
Add-Type -AssemblyName System.Data

class Forms {
    [System.Data.DataTable] CoverPage() {
        [Windows.Forms.MessageBox]::Show('SELECT THE COVER PAGE', '')

        $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
            InitialDirectory = [Environment]::GetFolderPath('Desktop')
            Title = "Select the cover page"
            Filter = "Word Documents (*.docx;*.doc)|*.docx;*.doc"
        }

        $word = New-Object -ComObject Word.Application
        $word.Visible = $true

        $coverPage_Table = New-Object System.Data.DataTable

        if ($FileBrowser.ShowDialog() -eq 'OK') {
            $wordDocumentPath = $FileBrowser.FileName
            $doc = $word.Documents.Open($wordDocumentPath)

            if ($doc.Tables.Count -ge 1) {
                $table = $doc.Tables.Item(1)

                for ($rowIndex = 1; $rowIndex -le $table.Rows.Count; $rowIndex++) {
                    $row = $table.Rows.Item($rowIndex)
                    $dataRow = $coverPage_Table.NewRow()

                    for ($colIndex = 1; $colIndex -le $row.Cells.Count; $colIndex++) {
                        $cell = $row.Cells.Item($colIndex)
                        $cellText = $cell.Range.Text.TrimEnd("`r", "`a")

                        if ($rowIndex -eq 1) {
                            $coverPage_Table.Columns.Add($cellText)
                        } else {
                            $dataRow[$colIndex - 1] = $cellText
                        }
                    }

                    if ($rowIndex -gt 1) {
                        $coverPage_Table.Rows.Add($dataRow)
                    }
                }
            }

            $doc.Close([Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges)
        }

        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)

        return $coverPage_Table
    }

    [string] CreditCheckEmail() {
        [Windows.Forms.MessageBox]::Show('SELECT THE CREDIT CHECK EMAIL', '')

        $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
            InitialDirectory = [Environment]::GetFolderPath('Desktop')
            Title = "Select the credit check"
            Filter = "Email files (*.msg;*.eml)|*.msg;*.eml|All files (*.*)|*.*"
        }

        $app = New-Object -ComObject Outlook.Application
        $out_mailBody = ""

        if ($FileBrowser.ShowDialog() -eq 'OK') {
            $emailFilePath = $FileBrowser.FileName
            $mailItem = $app.Session.OpenSharedItem($emailFilePath)
            $out_mailBody = $mailItem.Body
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($mailItem) | Out-Null
        }

        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null

        return $out_mailBody
    }

    [System.Data.DataTable[]] FinancialModel() {
        [Windows.Forms.MessageBox]::Show('SELECT THE FINANCIAL MODEL AS AN EXCEL FILE', '')
    
        $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
            InitialDirectory = [Environment]::GetFolderPath('Desktop')
            Title = "Select the Excel file"
            Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls"
        }
    
        $financialModelDT = New-Object System.Data.DataTable
        $financialModelDT3rdParty = New-Object System.Data.DataTable
        $workbook = $null
    
        if ($FileBrowser.ShowDialog() -eq 'OK') {
            $excelFilePath = $FileBrowser.FileName
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
    
            try {
                $workbook = $excel.Workbooks.Open($excelFilePath)
                $worksheet = $workbook.Sheets.Item("Equipment & Maint")
    
                $range = $worksheet.Range("A7").CurrentRegion
                $rowCount = $range.Rows.Count
                $colCount = $range.Columns.Count
    
                for ($col = 1; $col -le $colCount; $col++) {
                    $financialModelDT.Columns.Add()
                }
                
                for ($row = 1; $row -le $rowCount; $row++) {
                    $dataRow = $financialModelDT.NewRow()
                    for ($col = 1; $col -le $colCount; $col++) {
                        $dataRow[$col - 1] = $range.Item($row, $col).Text
                    }
                    $financialModelDT.Rows.Add($dataRow)
                }
                
                $financialModelDT.Rows | Where-Object {
                    $value = $_[0].ToString()
                    -not $value -or $value -in @("Recurring", "Country", "Charges", "Hardware", "Total", "Sub")
                } | ForEach-Object {
                    $financialModelDT.Rows.Remove($_)
                }
                
                $worksheet3rdParty = $workbook.Sheets.Item("3rd Party Services")
                $range3rdParty = $worksheet3rdParty.Range("A7").CurrentRegion
                $rowCount3rdParty = $range3rdParty.Rows.Count
                $colCount3rdParty = $range3rdParty.Columns.Count
                
                for ($col = 1; $col -le $colCount3rdParty; $col++) {
                    $financialModelDT3rdParty.Columns.Add()
                }
                
                for ($row = 1; $row -le $rowCount3rdParty; $row++) {
                    $dataRow = $financialModelDT3rdParty.NewRow()
                    for ($col = 1; $col -le $colCount3rdParty; $col++) {
                        $dataRow[$col - 1] = $range3rdParty.Item($row, $col).Text
                    }
                    $financialModelDT3rdParty.Rows.Add($dataRow)
                }
                
                $financialModelDT3rdParty.Rows | Where-Object {
                    $col0Value = $_[0].ToString()
                    $col1Value = $_[1].ToString()
                    -not $col0Value -or $col0Value -in @("Recurring", "Country") -or $col1Value -in @("AT&T", "AES", "Carousel")
                } | ForEach-Object {
                    $financialModelDT3rdParty.Rows.Remove($_)
                }
                
            } finally {
                if ($workbook) {
                    $workbook.Close($false)
                }
                $excel.Quit()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
            }
            
            return @($financialModelDT, $financialModelDT3rdParty)
        }
    
        return @($financialModelDT, $financialModelDT3rdParty)
    }
}



class RegexHandler {
    [string] $customerAddress
    [string] $customerContact
    [string] $salesAddress
    [string] $engagementManager

    RegexHandler([string] $customerAddress, [string] $customerContact, [string] $salesAddress, [string] $engagementManager) {
        $this.customerAddress = $customerAddress
        $this.customerContact = $customerContact
        $this.salesAddress = $salesAddress
        $this.engagementManager = $engagementManager
    }

    [string] GetCustomerStreet() {
        $customerStreet_edit = [System.Text.RegularExpressions.Regex]::Match($this.customerAddress, "Address:(.*?)City").Value.Trim()
        return [System.Text.RegularExpressions.Regex]::Replace($customerStreet_edit, "Address:|City", "").Trim()
    }

    [string] GetCustomerCity() {
        $customerCity_edit = [System.Text.RegularExpressions.Regex]::Match($this.customerAddress, "City:(.*?)State").Value.Trim()
        return [System.Text.RegularExpressions.Regex]::Replace($customerCity_edit, "City:|State", "").Trim()
    }

    [string] GetCustomerState() {
        return [System.Text.RegularExpressions.Regex]::Match($this.customerAddress, "AL|AK|AZ|AR|CA|CO|CT|DE|FL|GA|HI|ID|IL|IN|IA|KS|KY|LA|ME|MD|MA|MI|MN|MS|MO|MT|NE|NV|NH|NJ|NM|NY|NC|ND|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VT|VA|WA|WV|WI|WY").Value.Trim()
    }

    [string] GetCustomerZip() {
        return [System.Text.RegularExpressions.Regex]::Match($this.customerAddress, "\d{5}(-\d{4})?").Value.Trim()
    }

    [string] GetCustomerContactName() {
        $customerContactName_edit = [System.Text.RegularExpressions.Regex]::Match($this.customerContact, "Name:(.*?)Title").Value.Trim()
        return [System.Text.RegularExpressions.Regex]::Replace($customerContactName_edit, "Name:|Title", "").Trim()
    }

    [string] GetCustomerTitle() {
        $customerTitle_edit = [System.Text.RegularExpressions.Regex]::Match($this.customerContact, "Title:(.*?)Telephone").Value.Trim()
        return [System.Text.RegularExpressions.Regex]::Replace($customerTitle_edit, "Title:|Telephone", "").Trim()
    }

    [string] GetCustomerPhone() {
        $customerPhone_edit = [System.Text.RegularExpressions.Regex]::Match($this.customerContact, "Telephone:(.*?)Fax").Value.Trim()
        return [System.Text.RegularExpressions.Regex]::Replace($customerPhone_edit, "Telephone:|Fax", "").Trim()
    }

    [string] GetCustomerEmail() {
        return [System.Text.RegularExpressions.Regex]::Match($this.customerContact, "(?("")(""[^""]+?""@)|(([0-9a-zA-Z]((\.(?!\.))|[-!#$%&'*+/=?^_`{|}~\w])*)(?<=[0-9a-zA-Z])@))" + "((\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-zA-Z][-0-9a-zA-Z]*[0-9a-zA-Z]*\.)+[a-zA-Z0-9][\-a-zA-Z0-9]{0,22}[a-zA-Z0-9]))").Value.Trim()
    }

    [string] GetSalesStreet() {
        $salesStreet_edit = [System.Text.RegularExpressions.Regex]::Match($this.salesAddress, "Address:(.*?)City").Value.Trim()
        return [System.Text.RegularExpressions.Regex]::Replace($salesStreet_edit, "Address:|City", "").Trim()
    }

    [string] GetSalesCity() {
        $salesCity_edit = [System.Text.RegularExpressions.Regex]::Match($this.salesAddress, "City:(.*?)State").Value.Trim()
        return [System.Text.RegularExpressions.Regex]::Replace($salesCity_edit, "City:|State", "").Trim()
    }

    [string] GetSalesState() {
        return [System.Text.RegularExpressions.Regex]::Match($this.salesAddress, "AL|AK|AZ|AR|CA|CO|CT|DE|FL|GA|HI|ID|IL|IN|IA|KS|KY|LA|ME|MD|MA|MI|MN|MS|MO|MT|NE|NV|NH|NJ|NM|NY|NC|ND|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VT|VA|WA|WV|WI|WY").Value.Trim()
    }

    [string] GetSalesZip() {
        return [System.Text.RegularExpressions.Regex]::Match($this.salesAddress, "\d{5}(-\d{4})?").Value.Trim()
    }

    [string] GetSalesEmail() {
        $emailPattern = "(?("")(""[^""]+?""@)|(([0-9a-zA-Z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-zA-Z])@))" + "(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-zA-Z][-0-9a-zA-Z]*[0-9a-zA-Z]*\.)+[a-zA-Z0-9][\-a-zA-Z0-9]{0,22}[a-zA-Z0-9]))"
        return [System.Text.RegularExpressions.Regex]::Match($this.salesAddress, $emailPattern).Value.Trim()
    }

    [string] GetSalesManager() {
        $salesManager_edit = [System.Text.RegularExpressions.Regex]::Match($this.salesAddress, "Mgr:(.*?)SCVP").Value.Trim()
        return [System.Text.RegularExpressions.Regex]::Replace($salesManager_edit, "Mgr:|SCVP", "").Trim()
    }

    [string] GetSalesSCVP() {
        $salesSCVP_edit = [System.Text.RegularExpressions.Regex]::Match($this.salesAddress, "Name:\s*(.+)").Value.Trim()
        return [System.Text.RegularExpressions.Regex]::Replace($salesSCVP_edit, "Name:", "").Trim()
    }

    [string] GetEngagementManagerName() {
        $engagementManagerName_edit = [System.Text.RegularExpressions.Regex]::Match($this.engagementManager, "Name:(.*?)Address").Value.Trim()
        return [System.Text.RegularExpressions.Regex]::Replace($engagementManagerName_edit, "Name:|Address", "").Trim()
    }

    [string] GetEngagementManagerEmail() {
        $emailPattern = "(?("")(""[^""]+?""@)|(([0-9a-zA-Z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-zA-Z])@))" + "(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-zA-Z][-0-9a-zA-Z]*[0-9a-zA-Z]*\.)+[a-zA-Z0-9][\-a-zA-Z0-9]{0,22}[a-zA-Z0-9]))"
        return [System.Text.RegularExpressions.Regex]::Match($this.engagementManager, $emailPattern).Value.Trim()
    }
}

# Example usage:
$coverPage_Table = New-Object System.Data.DataTable
$coverPage_Table.Columns.Add("Column1")
$coverPage_Table.Columns.Add("Column2")
$coverPage_Table.Columns.Add("Column3")
$coverPage_Table.Rows.Add("John Doe", "", "Jane Doe")
$coverPage_Table.Rows.Add("", "", "")
$coverPage_Table.Rows.Add("Address:123 Main St City:Anytown State:CA 12345", "", "Address:456 Elm St City:Othertown State:NY 67890")
$coverPage_Table.Rows.Add("", "", "")
$coverPage_Table.Rows.Add("Name:John Smith Title:Manager Telephone:555-1234 Fax:555-5678", "", "Name:Jane Smith Title:Director Telephone:555-8765 Fax:555-4321")

$customerAddress = $coverPage_Table.Rows[2][0].ToString()
$customerContact = $coverPage_Table.Rows[4][0].ToString()
$salesAddress = $coverPage_Table.Rows[2][2].ToString()
$engagementManager = $coverPage_Table.Rows[4][2].ToString()

$regexHandler = [RegexHandler]::new($customerAddress, $customerContact, $salesAddress, $engagementManager)

$customerStreet = $regexHandler.GetCustomerStreet()
$customerCity = $regexHandler.GetCustomerCity()
$customerState = $regexHandler.GetCustomerState()
$customerZip = $regexHandler.GetCustomerZip()
$customerContactName = $regexHandler.GetCustomerContactName()
$customerTitle = $regexHandler.GetCustomerTitle()




class ExcelHandler {
    [void] $excel
    [void] $workbook
    [void] $worksheet

    ExcelHandler([string] $filePath, [string] $sheetName) {
        $this.excel = New-Object -ComObject Excel.Application
        $this.excel.Visible = $true
        $this.workbook = $this.excel.Workbooks.Open($filePath)
        $this.worksheet = $this.workbook.Sheets.Item($sheetName)
    }

    [void] SetCustomerInfo([string] $customerName, [string] $customerStreet, [string] $customerCity, [string] $customerState, [string] $customerZip, [string] $customerContactName, [string] $customerPhone, [string] $customerBillingCity, [string] $customerBillingState, [string] $customerBillingZip) {
        $this.worksheet.Cells.Item(12, 2).Value2 = $customerName  
        $this.worksheet.Cells.Item(13, 2).Value2 = $customerStreet  
        $this.worksheet.Cells.Item(14, 2).Value2 = $customerCity + ',' + $customerState + ',' + $customerZip
        $this.worksheet.Cells.Item(15, 2).Value2 = $customerContactName  
        $this.worksheet.Cells.Item(16, 2).Value2 = $customerPhone  
        $this.worksheet.Cells.Item(18, 2).Value2 = $customerName  
        $this.worksheet.Cells.Item(19, 2).Value2 = $customerName  
        $this.worksheet.Cells.Item(20, 2).Value2 = $customerName  
        $this.worksheet.Cells.Item(21, 2).Value2 = $customerBillingCity + ',' + $customerBillingState + ',' + $customerBillingZip  
        $this.worksheet.Cells.Item(22, 2).Value2 = $customerContactName  
        $this.worksheet.Cells.Item(23, 2).Value2 = $customerPhone  
    }

    [void] UpdateFinancialModel([System.Data.DataTable] $financialModelDT, [System.Data.DataTable] $financialModelDT3rdParty) {
        $rowNumber = 31
        if ($financialModelDT.Rows.Count -gt 0 -and $financialModelDT3rdParty.Rows.Count -gt 0) {
            foreach ($row in $financialModelDT.Rows) {
                $this.worksheet.Cells.Item($rowNumber, "A").Value2 = $row[2].ToString()
                $this.worksheet.Cells.Item($rowNumber, "B").Value2 = $row[3].ToString()
                $this.worksheet.Cells.Item($rowNumber, "E").Value2 = $row[4].ToString()
                $this.worksheet.Cells.Item($rowNumber, "F").Value2 = $row[7].ToString()
                $this.worksheet.Cells.Item($rowNumber, "I").Value2 = $row[14].ToString()
                $rowNumber++
            }

            $intStartingRow = 31 + $financialModelDT.Rows.Count
            $rowNumber += $intStartingRow - 31  

            foreach ($currentRow in $financialModelDT3rdParty.Rows) {
                $this.worksheet.Cells.Item($rowNumber, "A").Value2 = $currentRow[2].ToString()
                $this.worksheet.Cells.Item($rowNumber, "E").Value2 = $currentRow[3].ToString()
                $this.worksheet.Cells.Item($rowNumber, "I").Value2 = $currentRow[9].ToString()
                $rowNumber++
            }
        } elseif ($financialModelDT.Rows.Count -gt 0 -and $financialModelDT3rdParty.Rows.Count -eq 0) {
            foreach ($row in $financialModelDT.Rows) {
                $this.worksheet.Cells.Item($rowNumber, "A").Value2 = $row[2].ToString()
                $this.worksheet.Cells.Item($rowNumber, "B").Value2 = $row[3].ToString()
                $this.worksheet.Cells.Item($rowNumber, "E").Value2 = $row[4].ToString()
                $this.worksheet.Cells.Item($rowNumber, "F").Value2 = $row[7].ToString()
                $this.worksheet.Cells.Item($rowNumber, "I").Value2 = $row[14].ToString()
                $rowNumber++
            }
        } elseif ($financialModelDT.Rows.Count -eq 0 -and $financialModelDT3rdParty.Rows.Count -gt 0) {
            foreach ($row in $financialModelDT3rdParty.Rows) {
                $this.worksheet.Cells.Item($rowNumber, "A").Value2 = $row[2].ToString()
                $this.worksheet.Cells.Item($rowNumber, "E").Value2 = $row[3].ToString()
                $this.worksheet.Cells.Item($rowNumber, "I").Value2 = $row[9].ToString()
                $rowNumber++
            }
        }
    }

    [void] SaveWorkbook([string] $newFilePath) {
        $this.workbook.SaveAs($newFilePath)
    }

    [void] Close() {
        $this.workbook.Close($false)
        $this.excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($this.worksheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($this.workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($this.excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

# Example usage:
$filePath = "path_to_your_excel_file.xlsx"
$sheetName = "CPE ORDER"
$excelHandler = [ExcelHandler]::new($filePath, $sheetName)

$customerName = "John Doe"
$customerStreet = "123 Main St"
$customerCity = "Anytown"
$customerState = "CA"
$customerZip = "12345"
$customerContactName = "Jane Doe"
$customerPhone = "555-1234"
$customerBillingCity = "Anytown"
$customerBillingState = "CA"
$customerBillingZip = "12345"

$excelHandler.SetCustomerInfo($customerName, $customerStreet, $customerCity, $customerState, $customerZip, $customerContactName, $customerPhone, $customerBillingCity, $customerBillingState, $customerBillingZip)

# Assuming you have populated DataTables for financialModelDT and financialModelDT3rdParty
$financialModelDT = New-Object System.Data.DataTable
$financialModelDT3rdParty = New-Object System.Data.DataTable

$excelHandler.UpdateFinancialModel($financialModelDT, $financialModelDT3rdParty)

$newFilePath = "path_to_save_new_file.xlsx"
$excelHandler.SaveWorkbook($newFilePath)
$excelHandler.Close()

                # Example usage:
                # $forms = [Forms]::new()
                # $dataTable = $forms.CoverPage()
                # $dataTable | Format-Table -AutoSize
                # $emailBody = $forms.CreditCheckEmail()
                # Write-Output $emailBody
                # $financialModels = $forms.FinancialModel()
                # $financialModels[0] | Format-Table -AutoSize
                # $financialModels[1] | Format-Table -AutoSize
