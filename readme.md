ITG.Outlook
===========

Обёртки для COM интерфейсов MS Outlook для более комфортной "обработки" объектов пользователей и контактов AD, и не только.

Версия модуля: **1.0.3**

Функции модуля
--------------
			
### Contact
			
#### Get-Contact

Возвращаем найденный контакт из папки контактов по умолчанию.
	
	Get-Contact [-Filter <String>] [-WhatIf] [-Confirm] <CommonParameters>
	
	Get-Contact [-LastName <String>] [-FirstName <String>] [-MiddleName <String>] [-WhatIf] [-Confirm] <CommonParameters>
			
#### New-Contact

Создание нового контакта.
	
	New-Contact [-LastName <String>] [-FirstName <String>] [-MiddleName <String>] [-Initials <String>] [-NickName <String>] [-Subject <String>] [-Suffix <String>] [-Gender <String>] [-Birthday <DateTime>] [-Language <String>] [-Categories <String>] [-CompanyName <String>] [-Department <String>] [-JobTitle <String>] [-Profession <String>] [-AssistantName <String>] [-AssistantTelephoneNumber <String>] [-ManagerName <String>] [-BusinessAddressCity <String>] [-BusinessAddressState <String>] [-BusinessAddressCountry <String>] [-BusinessAddressPostalCode <String>] [-BusinessAddressStreet <String>] [-BusinessAddressPostOfficeBox <String>] [-OfficeLocation <String>] [-BusinessFaxNumber <String>] [-BusinessHomePage <String>] [-BusinessTelephoneNumber <String>] [-MobileTelephoneNumber <String>] [-Email1Address <String>] [-IMAddress <String>] [-InternetFreeBusyAddress <String>] [-PassThru] [-Force] [-WhatIf] [-Confirm] <CommonParameters>
			
#### Remove-Contact

Удаляем контакт.
	
	Remove-Contact [-InputObject] <__ComObject> [-WhatIf] [-Confirm] <CommonParameters>
			
#### Set-Contact

Редактируем реквизиты контакта.
	
	Set-Contact [-InputObject <__ComObject>] [-MAPIOBJECT <Object>] [-LastName <String>] [-FirstName <String>] [-MiddleName <String>] [-Initials <String>] [-NickName <String>] [-Subject <String>] [-Suffix <String>] [-Gender <String>] [-Birthday <DateTime>] [-Language <String>] [-Categories <String>] [-CompanyName <String>] [-Department <String>] [-JobTitle <String>] [-Profession <String>] [-AssistantName <String>] [-AssistantTelephoneNumber <String>] [-ManagerName <String>] [-BusinessAddressCity <String>] [-BusinessAddressState <String>] [-BusinessAddressCountry <String>] [-BusinessAddressPostalCode <String>] [-BusinessAddressStreet <String>] [-BusinessAddressPostOfficeBox <String>] [-OfficeLocation <String>] [-BusinessFaxNumber <String>] [-BusinessHomePage <String>] [-BusinessTelephoneNumber <String>] [-MobileTelephoneNumber <String>] [-Email1Address <String>] [-IMAddress <String>] [-InternetFreeBusyAddress <String>] [-PassThru] [-Force] [-WhatIf] [-Confirm] <CommonParameters>

Подробное описание функций модуля
---------------------------------
			
#### Get-Contact


Возвращаем найденный контакт из папки контактов по умолчанию.


##### Синтаксис
	
	Get-Contact [-Filter <String>] [-WhatIf] [-Confirm] <CommonParameters>
	
	Get-Contact [-LastName <String>] [-FirstName <String>] [-MiddleName <String>] [-WhatIf] [-Confirm] <CommonParameters>

##### Компонент

Outlook.Application

##### Параметры	

- `Filter <String>`
        Поисковый запрос в [синтаксисе Outlook][Синтаксис языка фильтров Outlook]
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?false
        Принимать подстановочные знаки?
        
- `LastName <String>`
        Фамилия
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `FirstName <String>`
        Имя
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `MiddleName <String>`
        Отчество
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `WhatIf [<SwitchParameter>]`
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?false
        Принимать подстановочные знаки?
        
- `Confirm [<SwitchParameter>]`
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?false
        Принимать подстановочные знаки?
        
- `<CommonParameters>`
        Данный командлет поддерживает общие параметры: Verbose, Debug,
        ErrorAction, ErrorVariable, WarningAction, WarningVariable,
        OutBuffer и OutVariable. Для получения дополнительных сведений введите
        "get-help about_commonparameters".





##### Примеры использования	

1. Пример 1.

		Get-Contact -Filter "[Subject]='Бетке Сергей Сергеевич'";

##### Связанные ссылки

- [Синтаксис языка фильтров Outlook]: http://office.microsoft.com/ru-ru/outlook-help/HA010238831.aspx
			
#### New-Contact

Создание нового контакта.


##### Синтаксис
	
	New-Contact [-LastName <String>] [-FirstName <String>] [-MiddleName <String>] [-Initials <String>] [-NickName <String>] [-Subject <String>] [-Suffix <String>] [-Gender <String>] [-Birthday <DateTime>] [-Language <String>] [-Categories <String>] [-CompanyName <String>] [-Department <String>] [-JobTitle <String>] [-Profession <String>] [-AssistantName <String>] [-AssistantTelephoneNumber <String>] [-ManagerName <String>] [-BusinessAddressCity <String>] [-BusinessAddressState <String>] [-BusinessAddressCountry <String>] [-BusinessAddressPostalCode <String>] [-BusinessAddressStreet <String>] [-BusinessAddressPostOfficeBox <String>] [-OfficeLocation <String>] [-BusinessFaxNumber <String>] [-BusinessHomePage <String>] [-BusinessTelephoneNumber <String>] [-MobileTelephoneNumber <String>] [-Email1Address <String>] [-IMAddress <String>] [-InternetFreeBusyAddress <String>] [-PassThru] [-Force] [-WhatIf] [-Confirm] <CommonParameters>

##### Компонент

Outlook.Application

##### Параметры	

- `LastName <String>`
        Фамилия
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `FirstName <String>`
        Имя
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `MiddleName <String>`
        Отчество
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `Initials <String>`
        Инициалы
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `NickName <String>`
        логин (он же - lname для почты и так далее)
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `Subject <String>`
        полное наименование контакта (используем ФИО)
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `Suffix <String>`
        мл. / ст. и так далее
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `Gender <String>`
        пол
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `Birthday <DateTime>`
        дата рождения
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `Language <String>`
        родной язык
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `Categories <String>`
        категории
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `CompanyName <String>`
        наименование компании
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `Department <String>`
        отдел
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `JobTitle <String>`
        должность
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `Profession <String>`
        профессия
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `AssistantName <String>`
        ФИО заместителя
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `AssistantTelephoneNumber <String>`
        телефон заместителя
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `ManagerName <String>`
        ФИО руководителя
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `BusinessAddressCity <String>`
        адрес - город
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `BusinessAddressState <String>`
        адрес - область
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `BusinessAddressCountry <String>`
        адрес - страна
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `BusinessAddressPostalCode <String>`
        адрес - индекс
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `BusinessAddressStreet <String>`
        адрес - улица и номер дома
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `BusinessAddressPostOfficeBox <String>`
        адрес - номер абонентского ящика
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `OfficeLocation <String>`
        адрес - номер кабинета
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `BusinessFaxNumber <String>`
        факс рабочий
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `BusinessHomePage <String>`
        сайт рабочий
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `BusinessTelephoneNumber <String>`
        телефон рабочий
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `MobileTelephoneNumber <String>`
        телефон мобильный
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `Email1Address <String>`
        адрес электронной почты
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `IMAddress <String>`
        # наименование контакта в списке контактов
        [Parameter(
            Mandatory=$false
            , ValueFromPipelineByPropertyName=$true
            , ParameterSetName="ContactProperties"
        )]
        [System.String]
        $Email1DisplayName = $Subject
        ,
        
        IM адрес
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `InternetFreeBusyAddress <String>`
        адрес, предоставляющий сведения о занятости
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `PassThru [<SwitchParameter>]`
        передавать домены далее по конвейеру или нет
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?false
        Принимать подстановочные знаки?
        
- `Force [<SwitchParameter>]`
        перезаписывать ли реквизиты существующих ящиков
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?false
        Принимать подстановочные знаки?
        
- `WhatIf [<SwitchParameter>]`
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?false
        Принимать подстановочные знаки?
        
- `Confirm [<SwitchParameter>]`
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?false
        Принимать подстановочные знаки?
        
- `<CommonParameters>`
        Данный командлет поддерживает общие параметры: Verbose, Debug,
        ErrorAction, ErrorVariable, WarningAction, WarningVariable,
        OutBuffer и OutVariable. Для получения дополнительных сведений введите
        "get-help about_commonparameters".





##### Примеры использования	

1. Пример 1.

		Get-Content $usersCsvFile | ConvertFrom-Csv -UseCulture | New-Contact;
			
#### Remove-Contact

Удаляем контакт.


##### Синтаксис
	
	Remove-Contact [-InputObject] <__ComObject> [-WhatIf] [-Confirm] <CommonParameters>

##### Компонент

Outlook.Application

##### Параметры	

- `InputObject <__ComObject>`
        Объект контакта, полученный через Get-Contact
        
        Требуется?                    true
        Позиция?                    1
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByValue)
        Принимать подстановочные знаки?
        
- `WhatIf [<SwitchParameter>]`
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?false
        Принимать подстановочные знаки?
        
- `Confirm [<SwitchParameter>]`
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?false
        Принимать подстановочные знаки?
        
- `<CommonParameters>`
        Данный командлет поддерживает общие параметры: Verbose, Debug,
        ErrorAction, ErrorVariable, WarningAction, WarningVariable,
        OutBuffer и OutVariable. Для получения дополнительных сведений введите
        "get-help about_commonparameters".





##### Примеры использования	

1. Пример 1.

		Get-Content $usersCsvFile | ConvertFrom-Csv -UseCulture | Get-Contact | Remove-Contact;
			
#### Set-Contact

Редактируем реквизиты контакта.


##### Синтаксис
	
	Set-Contact [-InputObject <__ComObject>] [-MAPIOBJECT <Object>] [-LastName <String>] [-FirstName <String>] [-MiddleName <String>] [-Initials <String>] [-NickName <String>] [-Subject <String>] [-Suffix <String>] [-Gender <String>] [-Birthday <DateTime>] [-Language <String>] [-Categories <String>] [-CompanyName <String>] [-Department <String>] [-JobTitle <String>] [-Profession <String>] [-AssistantName <String>] [-AssistantTelephoneNumber <String>] [-ManagerName <String>] [-BusinessAddressCity <String>] [-BusinessAddressState <String>] [-BusinessAddressCountry <String>] [-BusinessAddressPostalCode <String>] [-BusinessAddressStreet <String>] [-BusinessAddressPostOfficeBox <String>] [-OfficeLocation <String>] [-BusinessFaxNumber <String>] [-BusinessHomePage <String>] [-BusinessTelephoneNumber <String>] [-MobileTelephoneNumber <String>] [-Email1Address <String>] [-IMAddress <String>] [-InternetFreeBusyAddress <String>] [-PassThru] [-Force] [-WhatIf] [-Confirm] <CommonParameters>

##### Компонент

Outlook.Application

##### Параметры	

- `InputObject <__ComObject>`
        Объект контакта, полученный через Get-Contact
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByValue)
        Принимать подстановочные знаки?
        
- `MAPIOBJECT <Object>`
        Параметр, используется по сути только для определения переданного по конвейеру типа (потому как тип .net не опреде
        лить в этом случае)
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `LastName <String>`
        Фамилия
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `FirstName <String>`
        Имя
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `MiddleName <String>`
        Отчество
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `Initials <String>`
        Инициалы
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `NickName <String>`
        логин (он же - lname для почты и так далее)
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `Subject <String>`
        полное наименование контакта (используем ФИО)
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `Suffix <String>`
        мл. / ст. и так далее
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `Gender <String>`
        пол
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `Birthday <DateTime>`
        дата рождения
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `Language <String>`
        родной язык
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `Categories <String>`
        категории
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `CompanyName <String>`
        наименование компании
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `Department <String>`
        отдел
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `JobTitle <String>`
        должность
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `Profession <String>`
        профессия
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `AssistantName <String>`
        ФИО заместителя
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `AssistantTelephoneNumber <String>`
        телефон заместителя
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `ManagerName <String>`
        ФИО руководителя
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `BusinessAddressCity <String>`
        адрес - город
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `BusinessAddressState <String>`
        адрес - область
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `BusinessAddressCountry <String>`
        адрес - страна
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `BusinessAddressPostalCode <String>`
        адрес - индекс
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `BusinessAddressStreet <String>`
        адрес - улица и номер дома
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `BusinessAddressPostOfficeBox <String>`
        адрес - номер абонентского ящика
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `OfficeLocation <String>`
        адрес - номер кабинета
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `BusinessFaxNumber <String>`
        факс рабочий
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `BusinessHomePage <String>`
        сайт рабочий
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `BusinessTelephoneNumber <String>`
        телефон рабочий
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `MobileTelephoneNumber <String>`
        телефон мобильный
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `Email1Address <String>`
        адрес электронной почты
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `IMAddress <String>`
        # наименование контакта в списке контактов
        [Parameter(
            Mandatory=$false
            , ValueFromPipelineByPropertyName=$true
            , ParameterSetName="ContactProperties"
        )]
        [System.String]
        $Email1DisplayName = $Subject
        ,
        
        IM адрес
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `InternetFreeBusyAddress <String>`
        адрес, предоставляющий сведения о занятости
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?true (ByPropertyName)
        Принимать подстановочные знаки?
        
- `PassThru [<SwitchParameter>]`
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?false
        Принимать подстановочные знаки?
        
- `Force [<SwitchParameter>]`
        создание контакта в случае его отсутствия
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?false
        Принимать подстановочные знаки?
        
- `WhatIf [<SwitchParameter>]`
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?false
        Принимать подстановочные знаки?
        
- `Confirm [<SwitchParameter>]`
        
        Требуется?                    false
        Позиция?                    named
        Значение по умолчанию                
        Принимать входные данные конвейера?false
        Принимать подстановочные знаки?
        
- `<CommonParameters>`
        Данный командлет поддерживает общие параметры: Verbose, Debug,
        ErrorAction, ErrorVariable, WarningAction, WarningVariable,
        OutBuffer и OutVariable. Для получения дополнительных сведений введите
        "get-help about_commonparameters".





##### Примеры использования	

1. Пример 1.

		Get-Contact -Filter "[Subject]='Бетке Сергей Сергеевич'" | Set-Contact -Email1Address 'ivan.ivanov@domain.net';


