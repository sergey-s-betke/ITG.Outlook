[CmdletBinding(
	SupportsShouldProcess=$true,
	ConfirmImpact="Medium"
)]
param (
    [Parameter(
        Mandatory=$false,
        Position=0,
        ValueFromPipeline=$false,
        HelpMessage="Полный путь к файлу, из которого будем импортировать."
    )]
    [System.IO.FileInfo]
    $csvFile = `
        (join-path `
            -path ( ( [System.IO.FileInfo] ( $myinvocation.mycommand.path ) ).directory ) `
            -childPath 'users.csv' `
        )
)

Import-Module `
	(Join-Path `
		-Path ( Split-Path -Path ( $MyInvocation.MyCommand.Path ) -Parent ) `
        -ChildPath 'ITG.Outlook' `
    ) `
	-Force `
	-PassThru `
| Get-Readme -OutDefaultFile;

#get-content `
#    -path $csvFile `
#| convertFrom-csv `
#	-UseCulture `
#| New-Contact -Force -PassThru `
#| Select-Object Subject, EMail1DisplayName, EMail1Address `
#| Out-GridView `
#;
#