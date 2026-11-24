param(
  [string]$RosterPath = 'roster/cohort01-trainees.csv',
  [string]$TenantDomain = 'daraharsha7973gmail.onmicrosoft.com',
  [string]$CohortGroupName = 'SG-Cohort01-Trainees',
  [string]$TargetResourceGroup = 'rg-cohort01-lab'
)

$ErrorActionPreference = 'Stop'
$GraphToken = az account get-access-token --resource-type ms-graph --query accessToken -o tsv
$Headers = @{ Authorization = ('Bearer {0}' -f $GraphToken); 'Content-Type' = 'application/json' }

Write-Host ('Connected to Entra Tenant Domain: {0}' -f $TenantDomain) -ForegroundColor Cyan

# Check if any M365 License SKUs have available units
$Skus = (Invoke-RestMethod -Uri 'https://graph.microsoft.com/v1.0/subscribedSkus' -Headers $Headers).value
$AvailableSku = $Skus | Where-Object { ($_.prepaidUnits.enabled - $_.consumedUnits) -gt 0 } | Select-Object -First 1

# Ensure Cohort Security Group exists
$GroupUri = 'https://graph.microsoft.com/v1.0/groups?$filter=displayName eq ''{0}''' -f $CohortGroupName
$GroupResp = Invoke-RestMethod -Uri $GroupUri -Headers $Headers
if ($GroupResp.value.Count -eq 0) {
  $GroupBody = @{ displayName = $CohortGroupName; mailEnabled = $false; mailNickname = 'sgcohort01'; securityEnabled = $true; description = 'Automated security group for Cohort 01 trainees' } | ConvertTo-Json
  $CohortGroup = Invoke-RestMethod -Method Post -Uri 'https://graph.microsoft.com/v1.0/groups' -Headers $Headers -Body $GroupBody
} else {
  $CohortGroup = $GroupResp.value[0]
}
$GroupId = $CohortGroup.id

# Target Resource Group scope for Least-Privilege RBAC
$SubId = az account show --query id -o tsv
$RgExists = (az group exists --name $TargetResourceGroup) -eq 'true'
$RgScope = '/subscriptions/{0}/resourceGroups/{1}' -f $SubId, $TargetResourceGroup

$Roster = Import-Csv -Path $RosterPath
$AuditEntries = @()

foreach ($Row in $Roster) {
  $Upn = '{0}@{1}' -f $Row.MailNickname, $TenantDomain
  $UserUri = 'https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq ''{0}''' -f $Upn
  $User = (Invoke-RestMethod -Uri $UserUri -Headers $Headers).value | Select-Object -First 1

  if (-not $User) {
    Write-Warning ('User {0} not found in Entra directory; skipping.' -f $Upn)
    $AuditEntries += [PSCustomObject]@{ Timestamp = (Get-Date -Format 'o'); EmployeeId = $Row.EmployeeId; UPN = $Upn; Status = $Row.Status; Action = 'UserNotFound-Skipped'; Group = 'None'; License = 'None'; AzureRBAC = 'None' }
    continue
  }

  if ($Row.Status -eq 'Active') {
    # 1. Sync Entra Directory Attributes (Department & JobTitle)
    $PatchBody = @{ department = $Row.Department; jobTitle = $Row.JobTitle; usageLocation = 'IN' } | ConvertTo-Json
    Invoke-RestMethod -Method Patch -Uri ('https://graph.microsoft.com/v1.0/users/{0}' -f $User.id) -Headers $Headers -Body $PatchBody | Out-Null

    # 2. Add to Cohort Security Group (idempotent)
    $Members = (Invoke-RestMethod -Uri ('https://graph.microsoft.com/v1.0/groups/{0}/members' -f $GroupId) -Headers $Headers).value
    if ($Members.id -notcontains $User.id) {
      $MemberBody = @{ '@odata.id' = ('https://graph.microsoft.com/v1.0/directoryObjects/{0}' -f $User.id) } | ConvertTo-Json
      Invoke-RestMethod -Method Post -Uri ('https://graph.microsoft.com/v1.0/groups/{0}/members/$ref' -f $GroupId) -Headers $Headers -Body $MemberBody | Out-Null
    }

    # 3. Assign M365 License if available in tenant
    $LicenseStatus = 'Skipped-EntraFreeTenant'
    if ($AvailableSku) {
      $LicBody = @{ addLicenses = @(@{ skuId = $AvailableSku.skuId }); removeLicenses = @() } | ConvertTo-Json -Depth 4
      Invoke-RestMethod -Method Post -Uri ('https://graph.microsoft.com/v1.0/users/{0}/assignLicense' -f $User.id) -Headers $Headers -Body $LicBody | Out-Null
      $LicenseStatus = 'Assigned-{0}' -f $AvailableSku.skuPartNumber
    }

    # 4. Grant Least-Privilege Reader RBAC on Cohort Resource Group
    $RbacStatus = 'Deferred-CohortRGNotActive'
    if ($RgExists) {
      az role assignment create --assignee-object-id $User.id --assignee-principal-type User --role 'Reader' --scope $RgScope --output none
      $RbacStatus = 'Reader-Assigned'
    }

    $AuditEntries += [PSCustomObject]@{ Timestamp = (Get-Date -Format 'o'); EmployeeId = $Row.EmployeeId; UPN = $Upn; Status = 'Active'; Action = 'Onboarded-AttributesAndAccessSynced'; Group = $CohortGroupName; License = $LicenseStatus; AzureRBAC = $RbacStatus }
  }
  elseif ($Row.Status -eq 'Offboarded') {
    # 1. Tag Entra Directory Profile as Offboarded - Alumni
    $OffboardBody = @{ department = 'Offboarded - Alumni'; jobTitle = 'Former Trainee (Access Revoked)' } | ConvertTo-Json
    Invoke-RestMethod -Method Patch -Uri ('https://graph.microsoft.com/v1.0/users/{0}' -f $User.id) -Headers $Headers -Body $OffboardBody | Out-Null

    # 2. Remove from Cohort Security Group
    $Members = (Invoke-RestMethod -Uri ('https://graph.microsoft.com/v1.0/groups/{0}/members' -f $GroupId) -Headers $Headers).value
    if ($Members.id -contains $User.id) {
      Invoke-RestMethod -Method Delete -Uri ('https://graph.microsoft.com/v1.0/groups/{0}/members/{1}/$ref' -f $GroupId, $User.id) -Headers $Headers | Out-Null
    }

    # 3. Reclaim any assigned M365 licenses
    $UserLicenses = (Invoke-RestMethod -Uri ('https://graph.microsoft.com/v1.0/users/{0}/licenseDetails' -f $User.id) -Headers $Headers).value
    if ($UserLicenses.Count -gt 0) {
      $SkuIds = @($UserLicenses | ForEach-Object { $_.skuId })
      $RevokeBody = @{ addLicenses = @(); removeLicenses = $SkuIds } | ConvertTo-Json -Depth 4
      Invoke-RestMethod -Method Post -Uri ('https://graph.microsoft.com/v1.0/users/{0}/assignLicense' -f $User.id) -Headers $Headers -Body $RevokeBody | Out-Null
    }

    # 4. Strip Resource-Group-Scoped Azure RBAC Role Assignments
    if ($RgExists) {
      az role assignment delete --assignee $User.id --scope $RgScope --output none 2>$null
    }

    $AuditEntries += [PSCustomObject]@{ Timestamp = (Get-Date -Format 'o'); EmployeeId = $Row.EmployeeId; UPN = $Upn; Status = 'Offboarded'; Action = 'Offboarded-GroupAndRBACRevoked'; Group = 'Removed'; License = 'Reclaimed'; AzureRBAC = 'Revoked' }
  }
}

New-Item -ItemType Directory -Force -Path 'audit' | Out-Null
$AuditEntries | ConvertTo-Json -Depth 4 | Out-File -Encoding utf8 'audit/identity-sync-report.json'
$AuditEntries | Format-Table -AutoSize
