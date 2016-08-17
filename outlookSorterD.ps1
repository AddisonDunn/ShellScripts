# Addison Dunn, Michael Vabner 2016

# Helpful links
# http://stackoverflow.com/questions/22159170/grab-files-from-most-recently-received-email-in-specific-outlook-folder

# Get Outlook object
Add-type -assembly “Microsoft.Office.Interop.Outlook” | out-null
$outlookObj = new-object -comobject outlook.application
# Necessary line to grab namespace
$namespace = $outlookObj.GetNameSpace(“MAPI”)
# Get the collection of folders from the invoices account
$olFolders = $namespace.Folders.Item('invoices').Folders
# Get the inbox for the account
$inbox = $olFolders.Item('Inbox').Items | Sort-Object ReceivedTime -Descending

New-Item c:\users\svc.invoiceauto\documents\temp_folder -type directory -force
$main_filepath = “C:\Users\svc.invoiceauto\Documents\temp_folder” # Where we're gonna put all of the attachments

#Function to look through excel file and turn contents of first column into list 
$Excel = New-Object -ComObject Excel.Application 
$Excel.Visible = $true
$Excel.DisplayAlerts = $false
$ExcelWorkBook = $Excel.Workbooks.Open("C:\Users\svc.invoiceauto\Documents\FolderEmailListExceptions.xlsx") 
$ExcelWorkSheet = $Excel.Sheets.item("Sheet1") 
$ExcelWorkSheet.activate() 
$arrBlackListEmails = @()
$i = 1
Do 
{
    $arrBlackListEmails += $ExcelWorkSheet.Cells.Item($i, 1).Value()
    $i = $i + 1
}
Until ($ExcelWorkSheet.Cells.Item($i, 1).Value() -eq $null) # Move down until the last cell is empty
$excel.Quit()

# Loop through emails in inbox
for($i=1; $i -lt $inbox.Count; $i++)
{
    $email = $inbox.Item($i)

    # Check if there is an attachment and that the email is not checked
    If((0 -lt $email.Attachments.Count) -And  ( -Not $email.FlagStatus -eq 1))
    {
        # Get email address
        $address = $email.SenderEmailAddress
        # If email address is internal, this if-statement fixes the formatting
        If ($email.SenderEmailType -eq "EX") 
        {
            $address = $email.Sender.GetExchangeUser().PrimarySmtpAddress
        }

        # Get company name from email address
        $match = $address -replace ".*@" -replace ".com.*"
        $arrBlackListNames = @("Theresa Grouge", "Jen Dunlap", "Beverly Goodwin", "John Sankovich", "Sara Mallory", "Jacob Elliot")

        $b = $true # Boolean we'll use later

        # Check for the exceptions given by the excel file
        If( -Not $arrBlackListNames.Contains($email.SenderName) -And (-Not $arrBlackListEmails.Contains($address))) {
            Foreach ($element in $arrBlackListEmails)
            {
                If ($address -Match $element) # If the address matches one of the exceptions
                {
                    $b = $false
                                                                                                                                                                                                                                                                    
                    $filepath = $main_filepath + "\" + "MISCELLANEOUS" # Throw it in the misc folder

                    # This code is used further down and explained there.
                    If (-Not (Test-Path $filepath))
                    {
                        New-Item $filepath -type directory -force
                    }

                    $attachment = $email.Attachments.Item(1)
                    $x = 1
                    while( ($startingFilename -match "image00") )
                    {
                        $attachment = $email.Attachments.Item($x)
                        $x = $x + 1

                        If ($email.Attachments.Count -lt $x)
                        {
                            break
                        }
                    }

                    $startingFilename = $attachment.FileName
                    echo "MISC: $startingFilename"

                    $attachment | %{$_.saveasfile((join-path $filepath ($startingFilename)))}
                    echo "Loaded."
                    echo " "
                }
            }

        If ($b){

            # Load teh first attachment
            $attachment = $email.Attachments.Item(1)

            $filepath = $main_filepath + "\" + $match # Put the attachment in a folder with the company's name
            If (-Not (Test-Path $filepath))
            {
                New-Item $filepath -type directory -force
            }

            $date = $email.SentOn.ToString("yyyy-MM-dd") # We're going to put the date at the beginning of the filename

            $x = 1
            $b2 = $false
            $attachment = $email.Attachments.Item($x)
            $startingFilename = $attachment.FileName
            
            # InfoReliance emails often attach a logo whose filename starts with 'image00'. We don't
            # want that image, so we'll keep going until we find another kind of attachment.
            while( ($startingFilename -match "image00") ) 
            {
                $attachment = $email.Attachments.Item($x)
                $startingFilename = $attachment.FileName
                $x = $x + 1

                If ($email.Attachments.Count -lt $x)
                {
                    $b2 = $true
                    break
                }
            }
            If ( $b2 ) { continue } # The b2 variable is used to check whether the email only had the 'image00' files attached.

            echo "Starting filename: $startingFilename" # Output for testing purposes.
                
            $filename = $date + " " + $startingFilename

            $attachment | %{$_.saveasfile((join-path $filepath ($filename)))} # Saves the attachment
            echo "Loaded."
            echo " "

            $email.FlagStatus = 1 # Change the flagstatus to 1 to put a "check" next to the email.
               
                
                
        }
            
        }
        
       
    }
}