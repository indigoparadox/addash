#
# RemoteAD.ps1
# Some utility functions for working with the domain controller remotely.
# Requires admin credentials.
# Last Updated: indigoparadox, 2019/07/26
#

Function Invoke-OnDC {
    Param(
        [Parameter( Mandatory=$true, Position=0 )]
        [System.Management.Automation.PSCredential] $AdminCredential,
        [Parameter( Mandatory=$true, Position=1 )]
        [scriptblock] $ScriptBlock
    )
    
    $DomainController = Get-Registry -Hive HKCU -Path 'Software\ADDash' -Name 'DomainController'

    Return $(Invoke-Command -Credential $AdminCredential -ComputerName $DomainController `
        -ScriptBlock $ScriptBlock)
}

Function Get-RemoteADObject {
    Param(
        [string] $OU,
        [Parameter( Mandatory=$true )]
        [string] $ObjectType,
        [string] $Filter,
        [string] $GUID,
        [string] $RegistryPath,
        [System.Management.Automation.PSCredential] $AdminCredential
    )

    $Properties = "DistinguishedName","Enabled","Name","SamAccountName","SID","LockedOut"

    If( $ObjectType -eq 'User' ) {
	    $ADDObjects = Invoke-OnDC $AdminCredential { Get-ADUser -SearchBase $Using:OU -Filter $Using:Filter -Properties $Using:Properties } |
            Sort-Object -Property Name
    } ElseIf( $ObjectType -eq 'Computer' ) {
	    $ADDObjects = Invoke-OnDC $AdminCredential { Get-ADComputer -SearchBase $Using:OU -Filter $Using:Filter -Properties $Using:Properties } | 
            Sort-Object -Property Name
    } ElseIf( $ObjectType -eq 'Group' ) {
	    $ADDObjects = Invoke-OnDC $AdminCredential { Get-ADGroup -SearchBase $Using:OU -Filter $Using:Filter } |
            Sort-Object -Property Name
    } ElseIf( $ObjectType -eq 'Policies' ) {
	    $ADDObjects = Invoke-OnDC $AdminCredential { Get-GPO -All }
    } ElseIf( $ObjectType -eq 'PolicyReg' ) {
        Get-GPRegistryValue -Guid $Filter `
            -Key "HKCU\Software\Policies\Google\Chrome" `
            -ValueName "ManagedBookmarks" -ErrorAction SilentlyContinue
    }

    Return ,$ADDObjects
}
