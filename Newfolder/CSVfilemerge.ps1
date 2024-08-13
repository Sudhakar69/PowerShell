# Load the CSV data
$data = Import-Csv -Path "C:\Users\c-tgudise\OneDrive - APX\CVE.csv"

# Step 1: Update MachineDomain for specific Hostname
$data | ForEach-Object {
    if ($_.Hostname -eq "apxdev-usw-linuxbld") 
    {
        $_.MachineDomain = "apxna.com"
    }
}

# Step 2: Delete records with MachineDomain not equal to "apxna.com"
$data = $data | Where-Object { 
    $_.MachineDomain -eq "apxna.com" 
}


# Step 5: Sort by Severity and Hostname
$data = $data | Sort-Object -Property "HostnameDomain","Severity" | Select-Object -Unique

# Save the updated CSV data
$data | Export-Csv -Path "C:\Users\c-tgudise\OneDrive - APX\CVE_output.csv" -NoTypeInformation
