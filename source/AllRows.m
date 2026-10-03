let
    Columns = {"order_id","date","channel","sku","product","category","quantity","unit_price_cents","unit_cost_cents"},
    ReadCSV = (content as binary, expected as list) as table =>
        let t = Table.PromoteHeaders(Csv.Document(content,[Delimiter=",",Encoding=65001,QuoteStyle=QuoteStyle.Csv,ExtraValues=ExtraValues.Error]),[PromoteAllScalars=true])
        in if Table.ColumnNames(t) <> expected then error "CSV header differs from the contract" else t,
    Files = Table.SelectRows(Folder.Contents(DataRoot & "/monthly"),each [Extension] = ".csv"),
    Sorted = Table.Sort(Files,{{"Name",Order.Ascending}}),
    Loaded = List.Transform(Table.ToRecords(Sorted), (file) =>
        Table.AddColumn(Table.AddIndexColumn(ReadCSV(file[Content],Columns),"source_row",2,1,Int64.Type),"source_file",each file[Name],type text)),
    Raw = if List.IsEmpty(Loaded) then error "No monthly CSV files" else Table.Combine(Loaded),
    ProductRaw = ReadCSV(File.Contents(DataRoot & "/products.csv"),{"sku","product","category"}),
    Products = if Table.RowCount(ProductRaw)=0 or List.AnyTrue(List.Transform(ProductRaw[sku],each if _=null then true else Text.Trim(_)="")) or Table.RowCount(Table.Distinct(ProductRaw,{"sku"}))<>Table.RowCount(ProductRaw) then error "Product master has empty or repeated keys" else ProductRaw,
    Integer = (s as nullable text, positive as logical) as nullable number =>
        if s=null or s="" or Text.Select(s,{"0".."9"})<>s then null else
        let n=try Number.FromText(s,"en-US") otherwise null
        in if n=null then null else if n>=1000000000 or (positive and n=0) then null else Int64.From(n),
    Dates = Table.AddColumn(Raw,"DateTyped",each let d=try Date.FromText([date],[Format="yyyy-MM-dd",Culture="en-US"]) otherwise null in if d=null then null else if Date.ToText(d,"yyyy-MM-dd")<>[date] or Date.Year(d)<2020 or Date.Year(d)>2100 then null else d,type date),
    Qty = Table.AddColumn(Dates,"Qty",each Integer([quantity],true),Int64.Type),
    Price = Table.AddColumn(Qty,"PriceCents",each Integer([unit_price_cents],false),Int64.Type),
    Cost = Table.AddColumn(Price,"CostCents",each Integer([unit_cost_cents],false),Int64.Type),
    Counts = Table.Group(Cost,{"order_id"},{{"occurrences",each Table.RowCount(_),Int64.Type}}),
    WithCounts = Table.ExpandTableColumn(Table.NestedJoin(Cost,{"order_id"},Counts,{"order_id"},"counts",JoinKind.LeftOuter),"counts",{"occurrences"}),
    Joined = Table.NestedJoin(WithCounts,{"sku"},Products,{"sku"},"master",JoinKind.LeftOuter),
    Issues = Table.AddColumn(Joined,"issues",each Text.Combine(List.RemoveNulls({
        if [order_id]=null or Text.Trim([order_id])="" then "missing_order_id" else if [occurrences]>1 then "duplicate_order_id" else null,
        if [DateTyped]=null then "invalid_date" else null,
        if not List.Contains({"Online","Retail","Wholesale"},[channel]) then "invalid_channel" else null,
        if Table.RowCount([master])=0 then "unmatched_sku" else if [master]{0}[product]<>[product] or [master]{0}[category]<>[category] then "product_mismatch" else null,
        if [Qty]=null then "invalid_quantity" else null,
        if [PriceCents]=null then "invalid_unit_price_cents" else null,
        if [CostCents]=null then "invalid_unit_cost_cents" else null
    }),"; "),type text),
    Valued = Table.AddColumn(Issues,"valued",each [Qty]<>null and [PriceCents]<>null and [CostCents]<>null,type logical),
    Revenue = Table.AddColumn(Valued,"revenue_cents",each if [valued] then [Qty]*[PriceCents] else null,Int64.Type),
    Costs = Table.AddColumn(Revenue,"cost_cents",each if [valued] then [Qty]*[CostCents] else null,Int64.Type),
    Result = Table.RemoveColumns(Costs,{"master"})
in Result
