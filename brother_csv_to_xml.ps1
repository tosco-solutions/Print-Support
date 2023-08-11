##########################################
##CONVERT CSV to XML for Brother Devices##
##Paul Davies/Tosco Business & Education##
##########################################

## SETUP
## 1. You will need the csv file with the names and email addresses of the users that you wish to import.
## 2. The csv file will need three columns: id (a sequential list of numbers from 1 to x), name and email
## 3. You will need to alter the path to the csv file ($csvFilePath) (below). Your CSV should be saved in UTF-8 format.
## 4. You will need to alter the path to the xml file ($xmlFilePath) that you wish to output to (ensure the filename ends with .xml)
## 5. You will need to alter the brother model ($brotherModel) - (see procedure, below)

## PROCEDURE
## 1. Login to Web Management for the device using the device's IP Address in a web browser
## 2. Sign in to the device with Admin credentials and go to Address Book
## 3. The $brotherModel value is at the top of the screen (example: MFC-L6700DW series)
## 4. Click Export at the bottom left of the screen. Then click Export to file. Make a note of the file destinations, you'll need to use 'group.xml' when you re-import
## 5. Make sure you have followed the setup at the top of this script
## 6. Run the script from a powershell prompt: .\brother_csv_to_xml.ps1
## 7. Go back to Web Management, and click Import at the bottom left of the Address Book screen. For "Address Book" data file, click 'Browse' to the output.xml
## (or whatever you named it). For the "Group" data file, click 'Browse' and select the group.xml file that was exported in step 4.
## 8. Enjoy having your address book populated with the users and addresses

## NOTES and TODO
## Names are capped at 16 characters; if the user's name is longer than this it will be trimmed back to 16 characters.
## Names will be capitalised during Import to the machine. It's just how Brother rolls, I guess.
## ToDo: Remove the need for the ID column
## ToDo: Allow user to specify the column names


# Input and output file paths
#$csvFilePath = "C:\path\to\data.csv"
#$xmlFilePath = "C:\path\to\output.xml"
$csvFilePath = "C:\Users\pauld\Downloads\sample_addresses.csv"
$xmlFilePath = "C:\Users\pauld\Downloads\output.xml"
$brotherModel = "MFC-L6700DW series"

# Add XML header
$header = '<?xml version="1.0" encoding="UTF-8"?>
<vcards xmlns="urn:ietf:params:xml:ns:vcard-4.0"
    xmlns:ba="http://schemas.brother.info/mfc/controller/phx/2013/04/addressbookschemakeywords"
    ba:model="{0}" ba:dialkind="Speed">' -f $brotherModel

$footer = "</vcards>"

Add-Content -Path $xmlFilePath $header

# Open CSV file
$data = Import-Csv -Path $csvFilePath
# Function to cut user name to 16 characters
function LimitNameLength($name) {
    if ($name.Length -gt 16) {
        return $name.Substring(0, 16)
    }
    return $name
}

# Read each row of the CSV and get the id, name (cut to 16 characters), and email address then add them to an XML entry block
foreach ($row in $data) {
    $id = $row.id
    $name = LimitNameLength $row.name
    $email = $row.email
    $entry = '    <vcard ba:dial-id="{0}">
        <fn>
            <text>{1}</text>
        </fn>
        <email ba:index="1">
            <text>{2}</text>
        </email>
    </vcard>' -f $id, $name, $email
    Add-Content -Path $xmlFilePath $entry
}

# Add XML footer
Add-Content -Path $xmlFilePath $footer

# Job done.