let
    Month = Table.AddColumn(Accepted,"month",each Date.ToText([DateTyped],"yyyy-MM"),type text),
    Grouped = Table.Group(Month,{"month"},{{"orders",each Table.RowCount(_),Int64.Type},{"revenue_cents",each List.Sum([revenue_cents]),Int64.Type},{"cost_cents",each List.Sum([cost_cents]),Int64.Type}}),
    Revenue = Table.AddColumn(Grouped,"revenue_CAD",each [revenue_cents]/100,Currency.Type),
    Profit = Table.AddColumn(Revenue,"gross_profit_CAD",each ([revenue_cents]-[cost_cents])/100,Currency.Type),
    Margin = Table.AddColumn(Profit,"gross_margin",each if [revenue_cents]=0 then null else ([revenue_cents]-[cost_cents])/[revenue_cents],type number)
in Table.Sort(Margin,{{"month",Order.Ascending}})
