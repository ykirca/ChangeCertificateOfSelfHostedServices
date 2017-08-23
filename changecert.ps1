$serverListPath = "C:\Scripts\ssl_servers.csv" #Please change the path
$serverlist = Get-Content $serverListPath
$cred = Get-Credential 
foreach ($server in $serverlist)
{
    $server
    $session = New-PSSession -ComputerName $server -Credential $cred
    Invoke-Command -Session $session -ScriptBlock {new-item "C:\Users\administrator\Desktop\SSLCerts" -ItemType directory } #Create a directory on remote server named as SSLCerts 

    #Copy certificates from local computer to remote servers
    $destRootCA ="C:\Users\administrator\Desktop\SSLCerts\RootCA.cer" #Root certificate's name on remote computer. Please change the path.
    $destPrivate ="C:\Users\administrator\Desktop\SSLCerts\PrivateCert.pfx" #PFX certificate's name on remote computer. Please change the path.
    $localRootCA = 'C:\Scripts\RootCA.cer' #Root certificate on local computer. This certificate will be copied from local computer to remote server. Please change the path.
    $localPrivate = 'C:\Scripts\PrivateCert.pfx' #PFX certificate on local computer. This certificate will be copied from local computer to remote server. Please change the path.

    Copy-Item -Path $localRootCA -Destination $destRootCA -ToSession $session
    Copy-Item -Path $localPrivate -Destination $destPrivate -ToSession $session

    Invoke-Command -Session $session -ScriptBlock {
        Write-Host ("Working on: " + $server)

        Write-Output "Removing old certificate..."
        try
        {
            $certSubject = "CN=*.testdomain.com"
            $cert = Get-ChildItem cert:\LocalMachine\MY | Where-Object {$_.subject -like "$certSubject*" -AND $_.Subject -notmatch "CN=$env:COMPUTERNAME"}
            $oldthumbprint = $cert.Thumbprint.ToString()
            $removecert = Remove-Item -Path Cert:\localmachine\my\$oldthumbprint -DeleteKey
            Write-Output $removecert
            netsh http delete sslcert ipport=0.0.0.0:443

        }
        catch
        {
            Write-Output "An error occurred while removing old certificate!"
        }

     
        Write-Output "Adding new certificate..."
        try
        {
            Import-certificate -FilePath $destRootCA  -CertStoreLocation Cert:\LocalMachine\Root
            $mypwd = ConvertTo-SecureString -String "changeit" -Force -AsPlainText #Change password ("changeit")
            Import-PfxCertificate -FilePath $destPrivate cert:\localMachine\my -Password $mypwd 

            $certSubject = "CN=*.testdomain.com"
            $newcert = Get-ChildItem cert:\LocalMachine\MY | Where-Object {$_.subject -like "$certSubject*" -AND $_.Subject -notmatch "CN=$env:COMPUTERNAME"}
            $newthumbprint = $newcert.Thumbprint.ToString()

            netsh http add sslcert ipport=0.0.0.0:443 certhash=$newthumbprint appid='{06aabebd-3a91-4b80-8a15-adfd3c8a0b14}' certstore=my
            write-host $newthumbprint
            Write-Output "Completed!"
        }
        catch
        {
            Write-Output "An error occurred while adding new certificate!"
        }
    }

}
