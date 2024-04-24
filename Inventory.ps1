$ErrorActionPreference = 'Stop'

. "$PSScriptRoot\Config.ps1"
$certPassword = $Script:Config.CertificatePassword | ConvertTo-SecureString
$cert =  [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($Script:Config.CertificatePath, $certPassword)
Connect-MgGraph -ClientId $Script:Config.ClientId -Certificate $cert -TenantId $Script:Config.TenantId

$properties = @(
    'Id'
    'CreatedDateTime'
    'LicenseAssignmentStates'
    'AssignedLicenses'
    'OnPremisesImmutableId'
    'OnPremisesLastSyncDateTime'
    'OnPremisesProvisioningErrors'
    'OnPremisesSyncEnabled'
    'UserPrincipalName'
    'AccountEnabled'
    'UserType'
)

$userTable = New-Object -TypeName 'System.Data.DataTable' -ArgumentList @('EntraIDUser_stage')
[void]$userTable.Columns.Add('id', 'guid')
[void]$userTable.Columns.Add('createdDateTime', 'datetime')
[void]$userTable.Columns.Add('accountEnabled', 'boolean')
[void]$userTable.Columns.Add('onPremisesImmutableId', 'guid')
[void]$userTable.Columns.Add('onPremisesSyncEnabled', 'boolean')
[void]$userTable.Columns.Add('onPremisesLastSyncDateTime', 'datetime')
[void]$userTable.Columns.Add('onPremisesProvisioningErrors', 'boolean')
[void]$userTable.Columns.Add('userPrincipalName', 'string')

$licenseTable = New-Object -TypeName 'System.Data.DataTable' -ArgumentList @('MicrosoftLicenseAssigned_stage')
[void]$licenseTable.Columns.Add('id', 'int')
[void]$licenseTable.Columns.Add('userId', 'guid')
[void]$licenseTable.Columns.Add('skuId', 'guid')
[void]$licenseTable.Columns.Add('error', 'boolean')
[void]$licenseTable.Columns.Add('state', 'string')
[void]$licenseTable.Columns.Add('lastUpdatedDateTime', 'datetime')
[void]$licenseTable.Columns.Add('assignedByGroup', 'guid')
[void]$licenseTable.Columns.Add('assignedByGroupDisplayName', 'string')

function ValueOrNull($value) {
    if ($value) {
        $value
    }
    else {
        [System.DBNull]::Value
    }
}

function ToBool($value) {
    if ($value) {
        $true
    }
    else {
        $false
    }
}

$groupCache = @{}

foreach ($user in (Get-MgUser -Filter "userType eq 'Member'" -All -Property $properties)) {
    $onPremisesImmutableId = $null
    if ($user.OnPremisesImmutableId) {
        $bytes = [Convert]::FromBase64String($user.OnPremisesImmutableId)
        if ($bytes.Length -eq 16) {
            $onPremisesImmutableId = [Guid]::new($bytes)
        }
    }
    $row = $userTable.NewRow()
    $row['id'] = $user.Id
    $row['createdDateTime'] = $user.CreatedDateTime
    $row['accountEnabled'] = $user.AccountEnabled
    $row['onPremisesImmutableId'] = ValueOrNull $onPremisesImmutableId
    $row['onPremisesSyncEnabled'] = ToBool $user.OnPremisesSyncEnabled
    $row['onPremisesLastSyncDateTime'] = ValueOrNull $user.OnPremisesLastSyncDateTime
    $row['onPremisesProvisioningErrors'] = ($user.OnPremisesProvisioningErrors.Count -eq 0)
    $row['userPrincipalName'] = $user.UserPrincipalName
    $userTable.Rows.Add($row)
    
    foreach($license in $user.LicenseAssignmentStates) {
        $groupDisplayName = [System.DBNull]::Value
        if ($license.AssignedByGroup) {
            if ($groupCache.ContainsKey($license.AssignedByGroup)) {
                $groupDisplayName = $groupCache[$license.AssignedByGroup]
            }
            else {
                $groupDisplayName = (Get-MgGroup -GroupId $license.AssignedByGroup).DisplayName
                $groupCache[$license.AssignedByGroup] = $groupDisplayName
            }
        }
        $row = $licenseTable.NewRow()
        $row['userId'] = $user.Id
        $row['skuId'] = $license.SkuId
        $row['error'] = ($license.Error -ne 'None')
        $row['state'] = ValueOrNull $license.State
        $row['lastUpdatedDateTime'] = ValueOrNull $license.LastUpdatedDateTime
        $row['assignedByGroup'] = ValueOrNull $license.AssignedByGroup
        $row['assignedByGroupDisplayName'] = $groupDisplayName
        $licenseTable.Rows.Add($row)
    }
}

$conn = New-Object -TypeName 'System.Data.SqlClient.SqlConnection'
$conn.ConnectionString = $Script:Config.ConnectionString
$conn.Open()
$cmd = New-Object -TypeName 'System.Data.SqlClient.SqlCommand'
$cmd.Connection = $conn
$cmd.CommandType = 'Text'

$cmd.CommandText = 'TRUNCATE TABLE dbo.EntraIDUser_stage'
[void]$cmd.ExecuteNonQuery()

$cmd.CommandText = 'TRUNCATE TABLE dbo.MicrosoftLicenseAssigned_stage'
[void]$cmd.ExecuteNonQuery()

$bulkCopy = New-Object -TypeName 'System.Data.SqlClient.SqlBulkCopy' -ArgumentList @($conn)
$bulkCopy.DestinationTableName = 'EntraIDUser_stage'
$bulkCopy.WriteToServer($userTable)
$bulkCopy.Dispose()

$bulkCopy = New-Object -TypeName 'System.Data.SqlClient.SqlBulkCopy' -ArgumentList @($conn)
$bulkCopy.DestinationTableName = 'MicrosoftLicenseAssigned_stage'
$bulkCopy.WriteToServer($licenseTable)
$bulkCopy.Dispose()

$cmd.CommandText = 'dbo.spCommitMicrosoftOnline'
$cmd.CommandType = 'StoredProcedure'
[void]$cmd.ExecuteNonQuery()

$cmd.Dispose()


$conn.Dispose()
