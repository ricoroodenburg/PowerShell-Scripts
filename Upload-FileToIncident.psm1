Function Upload-FileToIncident {

    Param (

        [Parameter(Mandatory = $true)]
        [string]$URL,

        [Parameter(Mandatory = $true)]
        [string]$Username,

        [Parameter(Mandatory = $true)]
        [string]$ApplicationPassword,

        [Parameter(Mandatory = $true)]
        [string]$IncidentNumber,

        [Parameter(Mandatory = $true)]
        [string]$AttachmentPath


    )

    Begin {

        # Check if the file exists
        if (-not (Test-Path $AttachmentPath)) {
            Write-Error "Error: File not found at $AttachmentPath."
            return
        }

        # Convert the attachment file to base64
        $FileContent = [Convert]::ToBase64String((Get-Content -Path $AttachmentPath -Encoding Byte))
        
        $Headers = @{
        
            Authorization = "Basic " + [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($Username):$($ApplicationPassword)"))
            
        }

        $body = @"
--BOUNDARY
Content-Disposition: form-data; name="file"; filename="$($(Get-Item $AttachmentPath).Name)"
Content-Type: application/plain;charset=utf-8
Content-Transfer-Encoding: base64

$fileContent
--BOUNDARY--
"@

    }

    Process {

        Try {
            
            Invoke-RestMethod -Uri ($URL + "/tas/api/incidents/number/$IncidentNumber/attachments") -Headers $Headers -Method Post -Body $body -ContentType "multipart/form-data; boundary=BOUNDARY"
                
        } Catch {

            # Error...
            Write-Debug (">> Error while adding a new incident within TOPdesk, see error code below..")
            Write-Debug ("StatusCode:" + $_.Exception.Response.StatusCode.value__ )
            Write-Debug ("StatusDescription:" + $_.Exception.Response.StatusDescription)
            Write-Error $_
            Throw

        }

    }

}

Function Upload-FolderToTopdesk {

    Param (

        [Parameter(Mandatory = $true)]
        [string]$URL,

        [Parameter(Mandatory = $true)]
        [string]$Username,

        [Parameter(Mandatory = $true)]
        [string]$ApplicationPassword,

        [Parameter(Mandatory = $true)]
        [string]$IncidentNumber,

        [Parameter(Mandatory = $true)]
        [string]$FolderPath

    )

    Begin {

        # Check if folder exists
        if (-not (Test-Path $FolderPath)) {
            Write-Error "Error: Folder not found at $FolderPath."
            return
        }

        # Genereer pad voor tijdelijke zip
        $TempZipPath = "C:\Windows\Temp\TopdeskUpload_$((Get-Random)-join 6).zip"

        Try {
            # Maak de zip aan
            Compress-Archive -Path "$FolderPath\*" -DestinationPath $TempZipPath -Force
        }
        Catch {
            Write-Error "Fout bij het zippen van de folder: $_"
            return
        }

        # Check of zip succesvol is aangemaakt
        if (-not (Test-Path $TempZipPath)) {
            Write-Error "Zip-bestand is niet succesvol aangemaakt."
            return
        }

        # Lees bestand en converteer naar Base64
        $FileContent = [Convert]::ToBase64String((Get-Content -Path $TempZipPath -Encoding Byte))

        # Zet headers klaar
        $Headers = @{
            Authorization = "Basic " + [System.Convert]::ToBase64String(
                [System.Text.Encoding]::ASCII.GetBytes("$Username:$ApplicationPassword")
            )
        }

        # Maak multipart-form body
        $FileName = [IO.Path]::GetFileName($TempZipPath)
        $Body = @"
--BOUNDARY
Content-Disposition: form-data; name="file"; filename="$FileName"
Content-Type: application/zip
Content-Transfer-Encoding: base64

$FileContent
--BOUNDARY--
"@

    }

    Process {

        Try {
            Invoke-RestMethod -Uri "$URL/tas/api/incidents/number/$IncidentNumber/attachments" `
                -Headers $Headers `
                -Method Post `
                -Body $Body `
                -ContentType "multipart/form-data; boundary=BOUNDARY"

            Write-Output "✅ Upload succesvol uitgevoerd."

        }
        Catch {
            Write-Error "❌ Fout bij uploaden: $($_.Exception.Message)"
            if ($_.Exception.Response) {
                Write-Debug ("StatusCode: " + $_.Exception.Response.StatusCode.value__)
                Write-Debug ("StatusDescription: " + $_.Exception.Response.StatusDescription)
            }
            Throw
        }

    }

    End {
        # Opruimen
        if (Test-Path $TempZipPath) {
            Try {
                Remove-Item $TempZipPath -Force
                Write-Verbose "Tijdelijk bestand verwijderd: $TempZipPath"
            }
            Catch {
                Write-Warning "Kon het tijdelijke zip-bestand niet verwijderen: $_"
            }
        }
    }

}
