let

_version = "Excel-1.0.2",

// Max numbers of records API can return
_queryPageSize = 500,

_selectViewUIImpl = (url as text) as table =>
	let	
		tables = _getApiJson(url, "describe")[tables],
		tableSelector = #table(
			type table[ Table = text, Views = table],
			List.Transform(tables, (tentry) => { 
				tentry[recordName],
				(() as table => 
					let tdesc = _getApiJson(url, "describe", tentry[recordName]),
						viewSelector = #table(
							type table[ View = text, Data = table], 
							List.Transform(
								List.Select(tdesc[views], _isViewSupported),
								(vdesc) => { vdesc[name], _selectViewData(url, tdesc, vdesc) }))
						in
							_addTableKey(viewSelector, "View"))()
			}))
	in 
		_addTableKey(tableSelector, "Table"),

_selectViewImpl = (url as text, table as text, view as text, optional filter as text) =>
	let	tableChecked = if table = null then error "Table name is required" else table
	,	viewChecked = if view = null then error "View name is required" else view
	,	tableDesc = _getApiJson(url, "describe", tableChecked)
	,	viewDescs = List.Select(tableDesc[views], each Comparer.OrdinalIgnoreCase([name], viewChecked) = 0)
	,	viewDesc = if List.Count(viewDescs) <> 1 then error "View not found" else viewDescs{0}
	,	viewDescChecked = if _isViewSupported(viewDesc) then viewDesc else error Text.Format("Views of type ""#{0}"" are not supported", { viewDesc[type] })
	in 
		_selectViewData(url, tableDesc, viewDescChecked, filter),

_selectImpl = (url as text, table as text, columns as anynonnull, optional filter as text, optional sort as any) as any =>
	let query = []
	,	tableChecked = if table = null then error "Table name is required" else table
	,	columnList = if Value.Is(columns, type list) then columns else if Value.Is(columns, type text) then List.Transform(Text.Split(columns, ";"), Text.Trim) else error "Expected list or semicolon-separated text"
	,	columnListChecked = if List.Count(columnList) = 0 then error "At least one column is required" else columnList
	,	columnListCheckedStar = if List.Contains(columnListChecked, "*") then error "Star specifier is not allowed" else columnListChecked
	,	withColumns = Record.AddField(query, "column", columnList)
	,	withFilter = if filter <> null then Record.AddField(withColumns, "filter", filter) else withColumns
	,	withSort = if sort <> null then Record.AddField(withFilter, "sort", List.Transform(Text.Split(sort, ";"), Text.Trim)) else withFilter
	,	tableDesc = _getApiJson(url, "describe", tableChecked)
	,	columnDefTransform = _createColumnTranforms(tableDesc, columnListCheckedStar)
	,   resultList = _generateByPage((skip) => _selectGetPage(url, tableChecked, null, withSort, skip, _queryPageSize))
	in
		_createResult(resultList, columnDefTransform),
		
_selectViewData = (url as text, tableDesc as record, vdesc as record, optional filter as text) as table =>
	let query = []
	,	columnList = if vdesc[columns]? <> null then vdesc[columns] else List.Transform(List.Select(tableDesc[columns], each Text.Contains([displayOptions]?,"ShowInViews")), each [name])
	,	columnListWithGroups = if vdesc[groups]? <> null then vdesc[groups] & columnList else columnList
	,	columnDefTransform = _createColumnTranforms(tableDesc, columnListWithGroups)
	,	withFilter = if filter <> null then Record.AddField(query, "filter", filter) else query
	,   resultList = _generateByPage((skip) => _selectGetPage(url, tableDesc[recordName], vdesc[name], withFilter, skip, _queryPageSize))
	in
		_createResult(resultList, columnDefTransform),

_generateByPage = (getPage as function) as list =>
	List.Combine(
		List.Generate(
		() => getPage(0), 
		each _ <> null, 
		each if not [stop] then getPage([skip]) else null, 
		each [data])),

_selectGetPage = (url as text, table as text, view as nullable text, query as record, skip as number, pageSize as number) as record =>
	let withTop = if pageSize <> 500 then Record.AddField(query, "top", Text.From(pageSize)) else query
	,	withSkip = if skip <> 0 then Record.AddField(withTop, "skip", Text.From(skip)) else withTop
	,	rawData = _getApiJson(url, "select", table, view, withSkip)
	,	stop = List.Count(rawData) < pageSize
	,	data = List.Select(rawData, each _[#"@row.type"]? = null)
	in
		[ data = data, stop = stop, skip = skip + pageSize ],

_groupFuncs = { "EQ", "FW", "FL", "SS", "MI", "HH", "DD", "MM", "WK", "QQ", "YY", "DM", "MY", ".001", ".01", ".1", "1", "10", "100", "1K", "10K", "100K", "1M" },
_aggregateFuncs = { "COUNT", "STDEV", "STDEVP", "VAR", "VARP", "SUM", "AVG", "MIN", "MAX" },
_allSuffixes = List.Combine({ _groupFuncs, _aggregateFuncs }),

_stripSuffix = (value as text, suffixes as list) as text =>
	let nameParts = Text.Split(value, "//")
	, partCount = List.Count(nameParts)
	, suffix = if partCount > 1 then List.Last(nameParts) else ""
	in
		if suffix <> "" and List.Contains(suffixes, suffix) 
		then Text.Start(value, Text.Length(value) - 2 - Text.Length(suffix))
		else value,

_transformAggName = (value as text) as text =>
	let noSuffix = _stripSuffix(value, _aggregateFuncs)
	in
		if value <> noSuffix 
		then Text.Proper(Text.End(value, Text.Length(value) - Text.Length(noSuffix) - 2)) & " of " & noSuffix
		else value,

_getApiJson = (url as text, method as text, optional table as text, optional view as text, optional query as record) => 
	let apiURL = _getBaseApiURL(url) & 
		(if table <> null then "/" & _encodeUriPath(table) else "") & 
		(if view <> null then "/" & _encodeUriPath(view) else "") & 
		"/" & method & ".json"
	,	webResponse = Web.Contents(apiURL, [ Query = query, ManualStatusHandling = {400,403,404,405,409,500,503} ])
	,	responseStatus = Value.Metadata(webResponse)[Response.Status]
	,	jsonResponse = try Json.Document(webResponse, 65001) otherwise null
	in
		if jsonResponse <> null then
			if responseStatus = null or responseStatus >= 400 then 
				error Error.Record(
					"API Error " & Text.From(jsonResponse[error]) & "." & Text.From(jsonResponse[code]), 
					jsonResponse[message] & (if jsonResponse[source]? <> null then " (" & jsonResponse[source] & ")" else ""))
			else 
				jsonResponse
		else
			error "Unexpected result. Wrong URL?",

_getBaseApiURL = (url as text) =>
	let apiDelim = "/api/v2/"
	,	norm1 = if (try Int64.From(url) otherwise null) <> null then 
					"https://www.teamdesk.net/secure/api/v2/" & url 
				else Text.Replace(url, "/db/", apiDelim)
	,	dbId = Text.BeforeDelimiter(Text.AfterDelimiter(norm1, apiDelim), "/")
	in 
		Text.BeforeDelimiter(norm1, apiDelim) & apiDelim & dbId,

_encodeUriPath = (value as text) =>
	Uri.EscapeDataString(Text.Replace(Text.Replace(Text.Replace(Text.Replace(value, "%", "%25"), "/", "%2F"), "\", "%5C"), "?", "%3F")),

_createColumnTranforms = (tableDesc as record, columnList as list) as list =>
	let	columnListDistinct = List.Distinct(columnList)
	in
		List.Transform(columnListDistinct, (name) =>
			( let	noGroupSuffix = _stripSuffix(name, _groupFuncs)
				,	retainSuffix = name <> noGroupSuffix and List.Contains(columnList, noGroupSuffix)
				,	properName = if retainSuffix then name else _transformAggName(noGroupSuffix)
				,	noSuffixName = _stripSuffix(name, _allSuffixes)
				,	colDef = List.First(List.Select(tableDesc[columns], each Comparer.OrdinalIgnoreCase([name], noSuffixName) = 0), null)
			  in 
				if colDef = null then error Text.Format("API Error 400.3100: Column does not exist (#{0})", {name})
				// certain grouping and aggregate change type to numeric
				else if Text.EndsWith(name, "//DM") or 
				   Text.EndsWith(name, "//MY") or 
				   Text.EndsWith(name, "//COUNT") or 
				   colDef[type] = "Checkbox" and (Text.EndsWith(name, "//SUM") or Text.EndsWith(name, "//AVG"))
				   then { name, type number, (value) => value, properName }
				// group by YMDQW change timestamp to date
				else if colDef[type] = "Timestamp" and (Text.EndsWith(name, "//YY") or Text.EndsWith(name, "//MM") or Text.EndsWith(name, "//DD") or Text.EndsWith(name, "//WK") or Text.EndsWith(name, "//QQ"))
				   then { name, type date, (value) => Date.FromText(Text.Start(value, 10)), properName }
				else if colDef = null then { name, type text, (value) => Text.From(value), properName }
				else if colDef[type] = "Timestamp" then { name, type datetimezone, DateTimeZone.FromText, properName }
				else if colDef[type] = "Numeric" then { name, type number, (value) => value, properName }
				else if colDef[type] = "Checkbox" then { name, type logical, (value) => value, properName }
				else if colDef[type] = "Date" then { name, type date, (value) => Date.FromText(Text.Start(value, 10)), properName }
				else if colDef[type] = "Time" then { name, type time, (value) => Time.FromText(Text.Middle(value, 11, 8)), properName  }
				else if colDef[type] = "Duration" then { name, type duration, (value) => if value <> null then #duration(0, 0, 0, value) else null, properName }
				else { name, type text, (value) => value, properName })),

_createResult = (records as list, transforms as list) as table =>
	let	typedEmptyTable = Table.TransformColumnTypes(
		#table(List.Transform(transforms, (item) => item{3}), {}), 
		List.Transform(transforms, (item) => { item{3}, item{1} }))
	,	typedTableType = Value.Type(typedEmptyTable)
	,	recordTransform = List.Transform(transforms, (item) => { item{0}, item{2} })
	,	renameRecordFields = List.Transform(transforms, (item) => { item{0}, item{3} })
	,	typedListOfRecords = List.Transform(records, each Record.RenameFields(Record.TransformFields(_, recordTransform), renameRecordFields))
	in
		Table.FromRecords(typedListOfRecords, typedTableType),

_isViewSupported = (view as record) =>
	view[type] <> "SearchView" and view[type] <> "DashboardFilter" and view[type] <> "RecordPicker",

_addTableKey = (table as table, keyColumn as text) as table =>
	Value.ReplaceType(table, Type.AddTableKey(Value.Type(table), { keyColumn }, true)),

_selectImplMeta = Value.ReplaceType(_selectImpl, type function (
	url as (type text meta [
		Documentation.FieldCaption = "URL",
		Documentation.FieldDescription = "URL to your database",
		Documentation.SampleValues = { "https://www.teamdesk.net/secure/db/21995" }
	]),
	table as (type text meta [
		Documentation.FieldCaption = "Table",
		Documentation.FieldDescription = "Table name in singular form",
		Documentation.SampleValues = { "Invoice" }
	]), 
	columns as (type anynonnull meta [
		Documentation.FieldCaption = "Columns",
		Documentation.FieldDescription = "Semicolon-separated list of column names to retrieve",
		Documentation.SampleValues = { "Column1; Column2; Column3" }
	]), 
	optional filter as (type text meta [
		Documentation.FieldCaption = "Filter",
		Documentation.FieldDescription = "Filter expression in TeamDesk syntax",
		Documentation.SampleValues = { "not IsNull([Column])" }
	]), 
	optional sort as (type any meta [			
		Documentation.FieldCaption = "Sort by",
		Documentation.FieldDescription = "Semicolon-separated list of column names to sort by",
		Documentation.SampleValues = { "Column1; Column2//DESC" }
	])) as table meta [
		Documentation.Name = "Get the data from the table",
		Documentation.LongDescription = "Query table data by specifying table, columns to retrieve, filter and sort order"
	]),

_selectViewImplMeta = Value.ReplaceType(_selectViewImpl, type function (
	url as (type text meta [
		Documentation.FieldCaption = "URL",
		Documentation.FieldDescription = "URL to your database",
		Documentation.SampleValues = { "https://www.teamdesk.net/secure/db/21995" }
	]),
	table as (type text meta [
		Documentation.FieldCaption = "Table",
		Documentation.FieldDescription = "Table name in singular form",
		Documentation.SampleValues = { "Invoice" }
	]), 
	view as (type text meta [
		Documentation.FieldCaption = "View",
		Documentation.FieldDescription = "The name of the view",
		Documentation.SampleValues = { "List All" }
	]),
	optional filter as (type text meta [
		Documentation.FieldCaption = "Filter",
		Documentation.FieldDescription = "Filter expression in TeamDesk syntax",
		Documentation.SampleValues = { "not IsNull([Column])" }
	])) as table meta [
		Documentation.Name = "Get the data from the view",
		Documentation.LongDescription = "Query view data by specifying table and view name"
	]),

_selectViewUIImplMeta = Value.ReplaceType(_selectViewUIImpl, type function (
	url as (type text meta [
		Documentation.FieldCaption = "URL",
		Documentation.FieldDescription = "URL to your database",
		Documentation.SampleValues = { "https://www.teamdesk.net/secure/db/21995" }
	])) as table meta [
		Documentation.Name = "Browse database",
		Documentation.LongDescription = "Select the table and the view, get view's data"
	]),

TeamDesk.Select = _selectImplMeta,
TeamDesk.SelectView = _selectViewImplMeta,
TeamDesk.Database = _selectViewUIImplMeta

	in TeamDesk.Database