Add-Type `
	-AssemblyName 'Microsoft.Office.Interop.Outlook' `
;

function Get-Contact {
	<#
		.Component
			Outlook.Application
		.Synopsis
			Возвращаем найденный контакт из папки контактов по умолчанию.
		.Description
			Возвращаем найденный контакт из папки контактов по умолчанию.
		.Example
			Get-Contact -Filter "[Subject]='Бетке Сергей Сергеевич'";
	#>

	[CmdletBinding(
		SupportsShouldProcess=$true
		, ConfirmImpact="Low"
	)]
	
	param (
		# Поисковый запрос в синтаксисе Outlook
		[Parameter(
			Mandatory=$false
			, ParameterSetName="Filter"
		)]
		[String]
		[AllowNull()]
		$Filter
	,
		# Фамилия
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="Properties"
		)]
		[System.String]
		[ValidateNotNullOrEmpty()]
		[Alias("sn")]
		[Alias("SecondName")]
		$LastName
	,
		# Имя
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="Properties"
		)]
		[System.String]
		[ValidateNotNullOrEmpty()]
		[Alias("givenName")]
		$FirstName
	,
		# Отчество
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="Properties"
		)]
		[System.String]
		$MiddleName
	)
	
	begin {
		$Outlook = New-Object -ComObject Outlook.Application;
		$Contacts = $Outlook.GetNamespace('MAPI').GetDefaultFolder(
			[Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderContacts
		).Items;
	}
	process {
		if ( $PSCmdlet.ParameterSetName -ne 'Filter' ) {
			$Filter = (
				$PSBoundParameters.Keys `
				| ? { $PSBoundParameters.$_ } `
				| % { "[$_]='$($PSBoundParameters.$_)'" } `
			) -join ' AND ';
		};
		if ( $Filter ) {
			Write-Debug "Осуществляем выборку контактов по фильтру $Filter";
			for ( 
				$Contact = $Contacts.Find( $Filter );
				$Contact;
				$Contact = $Contacts.FindNext()
			) {
				return $Contact;
			};
		} else {
			$Contacts `
			| % {
				return $_;
			};
		};
	}
}  

function New-Contact {
	<#
		.Component
			Outlook.Application
		.Synopsis
			Создание нового контакта.
		.Description
			Создание нового контакта.
		.Example
			Get-Content `
			    -path $usersCsvFile `
			| ConvertFrom-Csv `
				-UseCulture `
			New-Contact;
	#>

	[CmdletBinding(
		SupportsShouldProcess=$true
		, ConfirmImpact="Medium"
	)]
	
	param (
		# Фамилия
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[ValidateNotNullOrEmpty()]
		[Alias("sn")]
		[Alias("SecondName")]
		$LastName
	,
		# Имя
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[ValidateNotNullOrEmpty()]
		[Alias("givenName")]
		$FirstName
	,
		# Отчество
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		$MiddleName
	,
		# Инициалы
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		$Initials
	,
		# логин (он же - lname для почты и так далее)
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[Alias("mailNickname")]
		$NickName
	,
		# полное наименование контакта (используем ФИО)
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[Alias("cn")]
		$Subject = ( ( $LastName, $FirstName, $MiddleName | ? { $_ } ) -join ' ' )
	,
		# мл. / ст. и так далее
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		$Suffix
<#
	,
		# г-н и так далее - обращение
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		$Title
#>
	,
		# пол
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[Alias("sex")]
		$Gender
	,
		# дата рождения
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.DateTime]
		$Birthday
	,
		# родной язык
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		$Language
	,
		# категории
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		$Categories
	,
		# наименование компании
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[Alias("company")]
		$CompanyName
	,
		# отдел
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		$Department
	,
		# должность
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[Alias("title")]
		$JobTitle
	,
		# профессия
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		$Profession
	,
		# ФИО заместителя
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[Alias("assistantCn")]
		$AssistantName
	,
		# телефон заместителя
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		$AssistantTelephoneNumber
	,
		# ФИО руководителя
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[Alias("managerCn")]
		$ManagerName
	,
		# адрес - город
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[Alias("l")]
		$BusinessAddressCity
	,
		# адрес - область
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[Alias("o")]
		$BusinessAddressState
	,
		# адрес - страна
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[Alias("co")]
		$BusinessAddressCountry
	,
		# адрес - индекс
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[Alias("postalCode")]
		$BusinessAddressPostalCode
	,
		# адрес - улица и номер дома
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[Alias("streetAddress")]
		$BusinessAddressStreet
	,
		# адрес - номер абонентского ящика
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		$BusinessAddressPostOfficeBox
	,
		# адрес - номер кабинета
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[Alias("room")]
		$OfficeLocation
	,
		# факс рабочий
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[Alias("facsimileTelephoneNumber")]
		$BusinessFaxNumber
	,
		# сайт рабочий
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[Alias("wWWHomePage")]
		$BusinessHomePage
	,
		# телефон рабочий
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[Alias("telephoneNumber")]
		$BusinessTelephoneNumber
	,
		# телефон мобильный
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[Alias("mobile")]
		$MobileTelephoneNumber
	,
		# адрес электронной почты
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[Alias("mail")]
		$Email1Address
	,
<#
		# наименование контакта в списке контактов
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		$Email1DisplayName = $Subject
	,
#>
		# IM адрес
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		$IMAddress = $Email1Address
	,
		# адрес, предоставляющий сведения о занятости
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		$InternetFreeBusyAddress
	,
		# передавать домены далее по конвейеру или нет
		[switch]
		$PassThru
	,
		# перезаписывать ли реквизиты существующих ящиков
		[switch]
		$Force
	)

	begin {
		$Outlook = New-Object -ComObject Outlook.Application;
		$Contacts = $Outlook.GetNamespace('MAPI').GetDefaultFolder(
			[Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderContacts
		).Items;
		$SetContact = ( { & (get-command Set-Contact) @PSBoundParameters } ).GetSteppablePipeline();
		$SetContact.Begin( $true );
	}
	process {
		$Params = @{};
		$GetContactParams = (Get-Command Get-Contact).Parameters;
		foreach ( $param in $PSBoundParameters.Keys ) {
			if ( $GetContactParams.ContainsKey($param) )  {
				$Params.$param = $PSBoundParameters.$param;
			};
		};
		$Contact = Get-Contact @Params;
		if ( $Contact -and -not $Force ) {
			Write-Error "Контакт $($Contact.Subject) существует. Для его перезаписи используйте ключ -Force.";
		} else {
			if ( -not $Contact ) {
				$Contact = $Contacts.Add( 2 );
			};
			$res = $SetContact.Process( $Contact );
			if ( $PassThru ) { return $Contact; }
		};
	}
	end {
		$SetContact.End();
	}
}  

function Set-Contact {
	<#
		.Component
			Outlook.Application
		.Synopsis
			Редактируем реквизиты контакта.
		.Description
			Редактируем реквизиты контакта.
		.Example
			Get-Contact -Filter "[Subject]='Бетке Сергей Сергеевич'" `
			| Set-Contact -Email1Address 'ivan.ivanov@domain.net' `
			;
	#>

	[CmdletBinding(
		SupportsShouldProcess=$true
		, ConfirmImpact="Medium"
	)]
	
	param (
		# Объект контакта, полученный через Get-Contact
		[Parameter(
			Mandatory=$false
			, ValueFromPipeline=$true
		)]
		[System.__ComObject]
		[Alias("Contact")]
		$InputObject
	,
		# Параметр, используется по сути только для определения переданного по конвейеру типа (потому как тип .net не определить в этом случае)
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
		)]
		$MAPIOBJECT
	,
		# Фамилия
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[ValidateNotNullOrEmpty()]
		[Alias("sn")]
		[Alias("SecondName")]
		$LastName
	,
		# Имя
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[ValidateNotNullOrEmpty()]
		[Alias("givenName")]
		$FirstName
	,
		# Отчество
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		$MiddleName
	,
		# Инициалы
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		$Initials
	,
		# логин (он же - lname для почты и так далее)
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[Alias("mailNickname")]
		$NickName
	,
		# полное наименование контакта (используем ФИО)
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[Alias("cn")]
		$Subject = ( ( $LastName, $FirstName, $MiddleName | ? { $_ } ) -join ' ' )
	,
		# мл. / ст. и так далее
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		$Suffix
<#
	,
		# г-н и так далее - обращение
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		$Title
#>
	,
		# пол
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[Alias("sex")]
		$Gender
	,
		# дата рождения
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.DateTime]
		$Birthday
	,
		# родной язык
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		$Language
	,
		# категории
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		$Categories
	,
		# наименование компании
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[Alias("company")]
		$CompanyName
	,
		# отдел
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		$Department
	,
		# должность
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[Alias("title")]
		$JobTitle
	,
		# профессия
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		$Profession
	,
		# ФИО заместителя
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[Alias("assistantCn")]
		$AssistantName
	,
		# телефон заместителя
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		$AssistantTelephoneNumber
	,
		# ФИО руководителя
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[Alias("managerCn")]
		$ManagerName
	,
		# адрес - город
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[Alias("l")]
		$BusinessAddressCity
	,
		# адрес - область
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[Alias("o")]
		$BusinessAddressState
	,
		# адрес - страна
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[Alias("co")]
		$BusinessAddressCountry
	,
		# адрес - индекс
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[Alias("postalCode")]
		$BusinessAddressPostalCode
	,
		# адрес - улица и номер дома
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[Alias("streetAddress")]
		$BusinessAddressStreet
	,
		# адрес - номер абонентского ящика
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		$BusinessAddressPostOfficeBox
	,
		# адрес - номер кабинета
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[Alias("room")]
		$OfficeLocation
	,
		# факс рабочий
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[Alias("facsimileTelephoneNumber")]
		$BusinessFaxNumber
	,
		# сайт рабочий
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[Alias("wWWHomePage")]
		$BusinessHomePage
	,
		# телефон рабочий
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[Alias("telephoneNumber")]
		$BusinessTelephoneNumber
	,
		# телефон мобильный
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[Alias("mobile")]
		$MobileTelephoneNumber
	,
		# адрес электронной почты
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		[Alias("mail")]
		$Email1Address
	,
<#
		# наименование контакта в списке контактов
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		$Email1DisplayName = $Subject
	,
#>
		# IM адрес
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		$IMAddress = $Email1Address
	,
		# адрес, предоставляющий сведения о занятости
		[Parameter(
			Mandatory=$false
			, ValueFromPipelineByPropertyName=$true
			, ParameterSetName="ContactProperties"
		)]
		[System.String]
		$InternetFreeBusyAddress
	,
		[switch]
		$PassThru
	,
		# создание контакта в случае его отсутствия
		[switch]
		$Force
	)

	begin {
		$Outlook = New-Object -ComObject Outlook.Application;
		$Contacts = $Outlook.GetNamespace('MAPI').GetDefaultFolder(
			[Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderContacts
		).Items;
	}
	process {
		if ( -not ( $InputObject -and $InputObject.MAPIOBJECT ) ) {
			$Params = @{};
			$GetContactParams = (Get-Command Get-Contact).Parameters;
			foreach ( $param in $PSBoundParameters.Keys ) {
				if ( $GetContactParams.ContainsKey($param) )  {
					$Params.$param = $PSBoundParameters.$param;
				};
			};
			$InputObject = Get-Contact @Params;
		};
		if ( -not ( $InputObject ) ) {
			if ( -not $Force ) {
				Write-Error "Контакт $($Contact.Subject) не существует. Для его создания используйте ключ -Force.";
			} else {
				if ( $PSCmdlet.ShouldProcess( $LastName, "Создание контакта" ) ) {
					Write-Verbose "Контакт для редактирования не передан явно и не обнаружен автоматически, создаём контакт.";
					$InputObject = $Contacts.Add( 2 );
				};
			};
		}
		if ( $InputObject ) {
			if ( $PSCmdlet.ShouldProcess( $InputObject.Subject, "Изменение реквизитов" ) ) {
				Write-Verbose "Изменяем реквизиты контакта $($InputObject.Subject).";
				$Params = (Get-Command Set-Contact).Parameters;
				$Params.Keys `
				| ? { $Params.$_.ParameterSets.ContainsKey('contactProperties') } `
				| ? { $PSBoundParameters.$_ } `
				| % {
					$InputObject.$_ = $PSBoundParameters.$_;
				} `
				;
				$InputObject.Save();
				$InputObject.Subject = ( 
					$InputObject.LastName, $InputObject.FirstName, $InputObject.MiddleName `
					| ? { $_ } 
				) -join ' ';
				if ( $Email1Address ) {
					$InputObject.Email1DisplayName = $InputObject.Subject;
				};
				$InputObject.Save();
				$InputObject.Close(0);
			};
		};
		if ( $PassThru ) { $input };
	}
}  

function Remove-Contact {
	<#
		.Component
			Outlook.Application
		.Synopsis
			Удаляем контакт
		.Description
			Удаляем контакт
		.Example
			Get-Content `
			    -path $usersCsvFile `
			| ConvertFrom-Csv `
				-UseCulture `
			| Get-Contact `
			| Remove-Contact `
			;
	#>

	[CmdletBinding(
		SupportsShouldProcess=$true
		, ConfirmImpact="High"
	)]
	
	param (
		# Объект контакта, полученный через Get-Contact
		[Parameter(
			Mandatory=$true
			, ValueFromPipeline=$true
		)]
		[System.__ComObject]
		[Alias("Contact")]
		$InputObject
	)

	process {
		if ( $PSCmdlet.ShouldProcess( $InputObject.Subject, "Удаление контакта" ) ) {
			Write-Verbose "Удаляем контакт $($InputObject.Subject).";
			$InputObject.Delete();
		};
	}
}  

Export-ModuleMember `
	Get-Contact `
	, New-Contact `
	, Set-Contact `
	, Remove-Contact `
;