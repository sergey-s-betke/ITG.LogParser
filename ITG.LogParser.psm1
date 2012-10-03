function Get-LPInputFormat{
    <#
        .Synopsis
            Возвращает объект - парсер журнала заданного типа.
        .Description
            Возвращает объект - парсер журнала заданного типа.
        .Parameter inputType
            Тип журнала, заданный строкой.
        .Example
            Поиск временных ошибок в журнале SMTP:
            Invoke-LPExecute ('
                "SELECT time, c-ip, cs-username, cs-method, cs-uri-stem, cs-uri-query, sc-status" + `
                " FROM $lastLogName" +`
                " WHERE cs-username='OutboundConnectionResponse' AND cs-uri-query LIKE '4%'", `
                Get-LPInputFormat("w3c") `
            )
       #>
 
    param (
        [Parameter(
            Mandatory=$true,
            Position=0,
            ValueFromPipeline=$false,
            HelpMessage="Тип журнала, заданный строкой."
        )]
        [string]$inputType
    )
 
    switch($inputType.ToLower()) {
        "ads"      {$inputobj = New-Object -comObject MSUtil.LogQuery.ADSInputFormat}
        "bin"      {$inputobj = New-Object -comObject MSUtil.LogQuery.IISBINInputFormat}
        "csv"      {$inputobj = New-Object -comObject MSUtil.LogQuery.CSVInputFormat}
        "etw"      {$inputobj = New-Object -comObject MSUtil.LogQuery.ETWInputFormat}
        "evt"      {$inputobj = New-Object -comObject MSUtil.LogQuery.EventLogInputFormat}
        "fs"       {$inputobj = New-Object -comObject MSUtil.LogQuery.FileSystemInputFormat}
        "httperr"  {$inputobj = New-Object -comObject MSUtil.LogQuery.HttpErrorInputFormat}
        "iis"      {$inputobj = New-Object -comObject MSUtil.LogQuery.IISIISInputFormat}
        "iisodbc"  {$inputobj = New-Object -comObject MSUtil.LogQuery.IISODBCInputFormat}
        "ncsa"     {$inputobj = New-Object -comObject MSUtil.LogQuery.IISNCSAInputFormat}
        "netmon"   {$inputobj = New-Object -comObject MSUtil.LogQuery.NetMonInputFormat}
        "reg"      {$inputobj = New-Object -comObject MSUtil.LogQuery.RegistryInputFormat}
        "textline" {$inputobj = New-Object -comObject MSUtil.LogQuery.TextLineInputFormat}
        "textword" {$inputobj = New-Object -comObject MSUtil.LogQuery.TextWordInputFormat}
        "tsv"      {$inputobj = New-Object -comObject MSUtil.LogQuery.TSVInputFormat}
        "urlscan"  {$inputobj = New-Object -comObject MSUtil.LogQuery.URLScanLogInputFormat}
        "w3c"      {$inputobj = New-Object -comObject MSUtil.LogQuery.W3CInputFormat}
        "xml"      {$inputobj = New-Object -comObject MSUtil.LogQuery.XMLInputFormat}
    }
     return $inputobj
}

function Invoke-LPExecute {
    <#
        .Synopsis
            Исполняет запрос через средство LogParser.
        .Description
            Представляет командлет для исполнения запросов через Log Parser. Возвращает объект RecordSet.
        .Parameter query
            Запрос в SQL синтаксисе Log Parser
        .Parameter inputType
            Интерфейс типизированного парсера для обрабатываемого журнала (создаваемый через Get-LPInputFormat).
            Если параметр не указан, тип журнала определяется Log Parser автоматически.
        .Example
            Поиск временных ошибок в журнале SMTP:
            Invoke-LPExecute ('
                "SELECT time, c-ip, cs-username, cs-method, cs-uri-stem, cs-uri-query, sc-status" + `
                " FROM $lastLogName" +`
                " WHERE cs-username='OutboundConnectionResponse' AND cs-uri-query LIKE '4%'" `
            )
       #>
 
    param (
        [Parameter(
            Mandatory=$true,
            Position=0,
            ValueFromPipeline=$false,
            HelpMessage="Запрос в SQL синтаксисе Log Parser."
        )]
        [string]$query,

        [Parameter(
            Mandatory=$false,
            Position=1,
            ValueFromPipeline=$false,
            HelpMessage="Интерфейс типизированного парсера для обрабатываемого журнала."
        )]
        $inputType
    )

    trap {
        $_.Exception.Data["extraInfo"] = "logParser query: `r`n$query";
        break;
    }
    $LPQuery = new-object -com MSUtil.LogQuery
    if($inputType) {
        $comInputType = Get-LPInputFormat $inputType;
        $LPRecordSet = $LPQuery.Execute($query, $comInputType);
    } else {
        $LPRecordSet = $LPQuery.Execute($query)
    }
    return $LPRecordSet
}

function Get-LPRecord {
    <#
        .Synopsis
            Возвращает PowerShell custom object из текущей записи Log Parser recordset.
        .Description
            Возвращает PowerShell custom object из текущей записи Log Parser recordset.
        .Parameter LPRecordSet
            RecordSet, результат Invoke-LPExecute
       #>
 
    param (
        [Parameter(
            Mandatory=$true,
            Position=0,
            ValueFromPipeline=$false,
            HelpMessage="RecordSet, результат Invoke-LPExecute."
        )]
        $LPRecordSet
    )

     $LPRecord = new-Object System.Management.Automation.PSObject;
    if (-not $LPRecordSet.atEnd()) {
        $Record = $LPRecordSet.getRecord();
        for ([int]$i = 0; $i -lt $LPRecordSet.getColumnCount(); $i++) {        
            $LPRecord | add-member `
                -memberType NoteProperty `
                -name $LPRecordSet.getColumnName($i) `
                -value $Record.getValue($i)
        }
    }
    return $LPRecord;
}
 
function Get-LPRecordSet {
    <#
        .Synopsis
            Исполняет запрос через средство LogParser.
        .Description
            Представляет командлет для исполнения запросов через Log Parser. Но возвращает уже массив объектов PS.
        .Parameter query
            Запрос в SQL синтаксисе Log Parser
        .Parameter inputType
            Интерфейс типизированного парсера для обрабатываемого журнала (создаваемый через Get-LPInputFormat).
            Если параметр не указан, тип журнала определяется Log Parser автоматически.
        .Example
            Поиск временных ошибок в журнале SMTP:
            Get-LPRecordSet ('
                "SELECT time, c-ip, cs-username, cs-method, cs-uri-stem, cs-uri-query, sc-status" + `
                " FROM $lastLogName" +`
                " WHERE cs-username='OutboundConnectionResponse' AND cs-uri-query LIKE '4%'" `
            )
       #>
 
    param (
        [Parameter(
            Mandatory=$true,
            Position=0,
            ValueFromPipeline=$false,
            HelpMessage="Запрос в SQL синтаксисе Log Parser."
        )]
        [string]$query,

        [Parameter(
            Mandatory=$false,
            Position=1,
            ValueFromPipeline=$false,
            HelpMessage="Интерфейс типизированного парсера для обрабатываемого журнала."
        )]
        $inputType
    )

    $LPRecordSet = Invoke-LPExecute -query $query -inputType $inputType;

    $LPRecords = new-object System.Management.Automation.PSObject[] 0;
    if (-not $LPRecordSet.atEnd()) {
        $exp = "select-object -property ```n" + (@( `
            for ([int]$i = 0; $i -lt $LPRecordSet.getColumnCount(); $i++) {
                "`t@{name=`"$($LPRecordSet.getColumnName($i))`";expression={`$_.getValue($i);};}";
            } `
        ) -join ", ```n");

        $LPRecords = @( for([int]$i=0; -not $LPRecordSet.atEnd(); $LPRecordSet.moveNext()) {
            #$LPRecords += Get-LPRecord($LPRecordSet)
            invoke-expression "`$LPRecordSet.getRecord() | $exp";
            $i++;
            if ($i -gt 20 ) {
                write-progress `
                    -id 1000 `
                    -activity "Загрузка записей журнала" `
                    -currentOperation "($i)" `
                    -status "Загрузка записей журнала"
            };
        });
        if ($i -gt 20 ) {
            write-progress `
                -id 1000 `
                -activity "Загрузка записей журнала" `
                -status "Загрузка записей журнала завершена. Загружено $($LPRecordSet.count) записей." `
                -completed
            };
    };
    $LPRecordSet.Close();
    return $LPRecords;
}

function Get-LPTableResults {
    <#
        .Synopsis
            Исполняет запрос через средство LogParser и возвращает таблицу типизированных данных.
        .Description
            Представляет командлет для исполнения запросов через Log Parser. Но возвращает уже таблицу типизированных данных PS.
        .Parameter query
            Запрос в SQL синтаксисе Log Parser
        .Parameter inputType
            Интерфейс типизированного парсера для обрабатываемого журнала (создаваемый через Get-LPInputFormat).
            Если параметр не указан, тип журнала определяется Log Parser автоматически.
        .Example
            Поиск временных ошибок в журнале SMTP:
            Get-LPTableResults ('
                "SELECT time, c-ip, cs-username, cs-method, cs-uri-stem, cs-uri-query, sc-status" + `
                " FROM $lastLogName" +`
                " WHERE cs-username='OutboundConnectionResponse' AND cs-uri-query LIKE '4%'" `
            )
       #>
 
    param (
        [Parameter(
            Mandatory=$true,
            Position=0,
            ValueFromPipeline=$false,
            HelpMessage="Запрос в SQL синтаксисе Log Parser."
        )]
        [string]$query,

        [Parameter(
            Mandatory=$false,
            Position=1,
            ValueFromPipeline=$false,
            HelpMessage="Интерфейс типизированного парсера для обрабатываемого журнала."
        )]
        $inputType
    )

    $LPRecordSet = Invoke-LPExecute -query $query -inputType $inputType;

    $tab = new-object System.Data.DataTable("Results")
    for($i = 0; $i -lt $LPRecordSet.getColumnCount(); $i++) {
        $col = new-object System.Data.DataColumn
        $ct = $LPRecordSet.getColumnType($i)
        $col.ColumnName = $LPRecordSet.getColumnName($i)

        switch ($LPRecordSet.getColumnType($i)) {
            "1"     { $col.DataType = [System.Type]::GetType("System.Int32") }
            "2"     { $col.DataType = [System.Type]::GetType("System.Double") }
            "4"     { $col.DataType = [System.Type]::GetType("System.DateTime") }
            default { $col.DataType = [System.Type]::GetType("System.String") }
        }

        $tab.Columns.Add($col)
    }

    for(; -not $LPRecordSet.atEnd(); $LPRecordSet.moveNext()) {
        $rowLP = $LPRecordSet.getRecord()
        $row = $tab.NewRow()
        for ($i = 0; ($i -lt $LPRecordSet.getColumnCount()); $i++) {
            $columnName = $LPRecordSet.getColumnName($i)
            $row.$columnName = $rowLP.getValue($i)
        }
        $tab.Rows.Add($row)
    }

    $ds = new-object System.Data.DataSet
    $ds.Tables.Add($tab)
    return $ds.Tables["Results"]
}

Export-ModuleMember `
    Get-LPInputFormat, `
    Invoke-LPExecute, `
    Get-LPRecord, `
    Get-LPRecordSet, `
    Get-LPTableResults
