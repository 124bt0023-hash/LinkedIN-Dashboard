# LinkedIN-Dashboard
# LinkedIn Company Page Dashboard — Power Query (M) + DAX reference



## 1. Power Query (M)

### Followers_Growth

```m
let
    Source = Excel.Workbook(File.Contents("C:\Users\YourName\Documents\LinkedIn_Company_Page_Dashboard_Data.xlsx"), null, true),
    Table = Source{[Item="FollowersGrowth", Kind="Table"]}[Data],
    ChangedType = Table.TransformColumnTypes(Table, {
        {"Date", type date},
        {"New Followers (Organic)", Int64.Type},
        {"New Followers (Paid)", Int64.Type},
        {"New Followers (Total)", Int64.Type},
        {"Total Followers", Int64.Type}
    })
in
    ChangedType
```

### Posts_Performance

```m
let
    Source = Excel.Workbook(File.Contents("C:\Users\YourName\Documents\LinkedIn_Company_Page_Dashboard_Data.xlsx"), null, true),
    Table = Source{[Item="PostsPerformance", Kind="Table"]}[Data],
    ChangedType = Table.TransformColumnTypes(Table, {
        {"Post ID", type text},
        {"Date", type date},
        {"Content Type", type text},
        {"Campaign", type text},
        {"Impressions", Int64.Type},
        {"Clicks", Int64.Type},
        {"Reactions", Int64.Type},
        {"Comments", Int64.Type},
        {"Shares", Int64.Type},
        {"Engagement Rate", type number},
        {"CTR", type number},
        {"Total Engagements", Int64.Type}
    })
in
    ChangedType
```

### Page_Visitors

```m
let
    Source = Excel.Workbook(File.Contents("C:\Users\YourName\Documents\LinkedIn_Company_Page_Dashboard_Data.xlsx"), null, true),
    Table = Source{[Item="PageVisitors", Kind="Table"]}[Data],
    ChangedType = Table.TransformColumnTypes(Table, {
        {"Date", type date},
        {"Page Views", Int64.Type},
        {"Unique Visitors", Int64.Type},
        {"Desktop Views", Int64.Type},
        {"Mobile Views", Int64.Type}
    })
in
    ChangedType
```

### Demographics

```m
let
    Source = Excel.Workbook(File.Contents("C:\Users\YourName\Documents\LinkedIn_Company_Page_Dashboard_Data.xlsx"), null, true),
    Table = Source{[Item="Demographics", Kind="Table"]}[Data],
    ChangedType = Table.TransformColumnTypes(Table, {
        {"Category", type text},
        {"Segment", type text},
        {"Percentage", Int64.Type}
    })
in
    ChangedType
```

### Traffic_Sources

```m
let
    Source = Excel.Workbook(File.Contents("C:\Users\YourName\Documents\LinkedIn_Company_Page_Dashboard_Data.xlsx"), null, true),
    Table = Source{[Item="TrafficSources", Kind="Table"]}[Data],
    ChangedType = Table.TransformColumnTypes(Table, {
        {"Month", type date},
        {"Source", type text},
        {"Sessions", Int64.Type}
    })
in
    ChangedType
```

### KPI_Monthly

```m
let
    Source = Excel.Workbook(File.Contents("C:\Users\YourName\Documents\LinkedIn_Company_Page_Dashboard_Data.xlsx"), null, true),
    Table = Source{[Item="KPIMonthly", Kind="Table"]}[Data],
    ChangedType = Table.TransformColumnTypes(Table, {
        {"Month", type date},
        {"New Followers", Int64.Type},
        {"Total Page Views", Int64.Type},
        {"Posts Published", Int64.Type},
        {"Total Impressions", Int64.Type},
        {"Total Engagements", Int64.Type},
        {"Avg Engagement Rate", type number}
    })
in
    ChangedType
```

---

## 2. Date table (DAX calculated table)


DateTable =
ADDCOLUMNS (
    CALENDAR ( MIN ( Followers_Growth[Date] ), MAX ( Followers_Growth[Date] ) ),
    "Year", YEAR ( [Date] ),
    "MonthNumber", MONTH ( [Date] ),
    "MonthName", FORMAT ( [Date], "MMM" ),
    "MonthYear", FORMAT ( [Date], "MMM YYYY" ),
    "Quarter", "Q" & FORMAT ( [Date], "Q" ),
    "WeekdayName", FORMAT ( [Date], "ddd" ),
    "IsWeekend", IF ( WEEKDAY ( [Date], 2 ) > 5, TRUE, FALSE )
)
```

**Relationships to build in Model view (all single-direction, 1 → many):**

| From | To |
|---|---|
| `DateTable[Date]` | `Followers_Growth[Date]` |
| `DateTable[Date]` | `Posts_Performance[Date]` |
| `DateTable[Date]` | `Page_Visitors[Date]` |
| `DateTable[Date]` | `Traffic_Sources[Month]` |
| `DateTable[Date]` | `KPI_Monthly[Month]` |

---

## 3. DAX measures

Create a blank table named `_Measures` (New table → `_Measures = ROW("x", 0)`,
then hide the `x` column) and add all measures there — keeps them out of your
data tables in the field list.

### Follower measures

```dax
Total Followers =
MAX ( Followers_Growth[Total Followers] )
```

```dax
New Followers (Period) =
SUM ( Followers_Growth[New Followers (Total)] )
```

```dax
New Followers (Organic) =
SUM ( Followers_Growth[New Followers (Organic)] )
```

```dax
New Followers (Paid) =
SUM ( Followers_Growth[New Followers (Paid)] )
```

```dax
Follower Growth MoM % =
VAR CurrentFollowers = [Total Followers]
VAR PriorFollowers =
    CALCULATE (
        [Total Followers],
        DATEADD ( DateTable[Date], -1, MONTH )
    )
RETURN
    DIVIDE ( CurrentFollowers - PriorFollowers, PriorFollowers )
```

```dax
Net New Followers YTD =
TOTALYTD ( [New Followers (Period)], DateTable[Date] )
```

### Content / engagement measures

```dax
Total Impressions =
SUM ( Posts_Performance[Impressions] )
```

```dax
Total Engagements =
SUM ( Posts_Performance[Total Engagements] )
```

```dax
Total Clicks =
SUM ( Posts_Performance[Clicks] )
```

```dax
Total Reactions =
SUM ( Posts_Performance[Reactions] )
```

```dax
Engagement Rate =
DIVIDE ( [Total Engagements], [Total Impressions] )
```

```dax
Avg CTR =
AVERAGE ( Posts_Performance[CTR] )
```

```dax
Posts Published =
DISTINCTCOUNT ( Posts_Performance[Post ID] )
```

```dax
Avg Engagements per Post =
DIVIDE ( [Total Engagements], [Posts Published] )
```

```dax
Best Performing Content Type =
VAR RankedTypes =
    TOPN (
        1,
        SUMMARIZE ( Posts_Performance, Posts_Performance[Content Type], "Eng", [Total Engagements] ),
        [Eng], DESC
    )
RETURN
    MAXX ( RankedTypes, Posts_Performance[Content Type] )
```

### Page visitor / traffic measures

```dax
Total Page Views =
SUM ( Page_Visitors[Page Views] )
```

```dax
Total Unique Visitors =
SUM ( Page_Visitors[Unique Visitors] )
```

```dax
Mobile Views % =
DIVIDE ( SUM ( Page_Visitors[Mobile Views] ), [Total Page Views] )
```

```dax
Desktop Views % =
DIVIDE ( SUM ( Page_Visitors[Desktop Views] ), [Total Page Views] )
```

```dax
Total Sessions (Traffic Sources) =
SUM ( Traffic_Sources[Sessions] )
```

```dax
Top Traffic Source =
VAR RankedSources =
    TOPN (
        1,
        SUMMARIZE ( Traffic_Sources, Traffic_Sources[Source], "S", [Total Sessions (Traffic Sources)] ),
        [S], DESC
    )
RETURN
    MAXX ( RankedSources, Traffic_Sources[Source] )
```

### Demographics helper measure

```dax
Demographic % =
AVERAGE ( Demographics[Percentage] ) / 100
```

## 4. Suggested visual → measure mapping

| Visual | Measures / fields |
|---|---|
| KPI cards (Overview page) | `Total Followers`, `Total Impressions`, `Engagement Rate`, `Total Page Views` |
| Follower growth line chart | `DateTable[MonthYear]` (axis) × `New Followers (Period)` |
| Engagement by content type (bar) | `Posts_Performance[Content Type]` (axis) × `Total Engagements` |
| Top 10 posts (table) | `Posts_Performance` columns, sorted by `Total Engagements` desc, `TOPN` filter |
| Audience donut/bar charts | `Demographics[Segment]` sliced by `Demographics[Category]`, values = `Demographic %` |
| Traffic by source (stacked column) | `DateTable[MonthYear]` (axis) × `Traffic_Sources[Source]` (legend) × `Total Sessions (Traffic Sources)` |
| Desktop vs. mobile (donut) | `Desktop Views %`, `Mobile Views %` |
