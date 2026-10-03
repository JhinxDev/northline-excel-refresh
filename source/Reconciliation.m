let
    Sum = (xs as list) as number => List.Sum({0} & List.RemoveNulls(xs)),
    RawRows = Table.RowCount(AllRows),
    AcceptedRows = Table.RowCount(Accepted),
    ExceptionRows = Table.RowCount(Exceptions),
    RawValue = Sum(AllRows[revenue_cents]),
    AcceptedValue = Sum(Accepted[revenue_cents]),
    ExcludedValue = Sum(Exceptions[revenue_cents]),
    Result = #table(type table [metric=text,value=Int64.Type],{
        {"input_rows",RawRows},{"accepted_rows",AcceptedRows},{"exception_rows",ExceptionRows},
        {"unvalued_rows",Table.RowCount(Table.SelectRows(AllRows,each not [valued]))},
        {"raw_valued_revenue_cents",RawValue},{"excluded_valued_revenue_cents",ExcludedValue},
        {"revenue_cents",AcceptedValue},{"cost_cents",Sum(Accepted[cost_cents])},
        {"row_difference",RawRows-AcceptedRows-ExceptionRows},{"revenue_difference_cents",RawValue-AcceptedValue-ExcludedValue}
    })
in Result
