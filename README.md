# XTrack-xml-to-mdd

## METADATA 
- All languages must have the following lists integrated, exactly as below:
```vb
CATEGORIES_LIST "" define
{
    _1 "cat1",
    _2 "cat2"
};
BRANDLIST_TEXT_ONLY "" define
{
    _1 "brand1",
    _2 "brand2"
};
BRANDLIST_LOGOS "" define
{
    _1 "brand1",
    _2 "brand2"
};
BRANDLIST_CLOSENESS "" define
{
    _1 "brand1",
    _2 "brand2"
};
BRANDLIST_TEXT_LOGO "" define
{
    _1 "brand1",
    _2 "brand2"
};
STATEMENT_LIST "" define
{
    _1 "statement1",
    _2 "statement2"
};
TOUCHPOINT_LIST "" define
{
    _1 "touchpoint1",
    _2 "touchpoint2"
};

```
- All questions that use the lists above must be updated with the new ones.
- The questions that contain Statement Lists must contain the lists in the following format:
    STATEMENT_LIST_<span style="background-color: blue">QUESTIONNAME</span>
    - Ex: STATEMENT_LIST_<span style="background-color: blue">BIA</span>

## ROUTING
- Delete all dims used for the brand image insert (dim brand1,brand2, …)
- Before the first question that uses CATEGORIES_LIST add the following placeholders:
```vb
'INSERT CATEGORY FILTER BELOW

'INSERT CATEGORY FILTER ABOVE
```
- Before the first question that uses a BRANDLIST add the following placeholders:
```vb
'INSERT BRAND FILTER AND IMAGES BELOW

'INSERT BRAND FILTER AND IMAGES ABOVE
```
- Before the first question that uses a STATEMENTLIST add the following placeholders:
```vb
'INSERT STATEMENTS FILTER BELOW

'INSERT STATEMENTS FILTER ABOVE
```
- Before the first question that uses TOUCHPOINT_LIST add the following placeholders:
```vb
'INSERT TOUCHPOINTS FILTER BELOW

'INSERT TOUCHPOINTS FILTER ABOVE
```
- Update the filters for the questions that need categories filtering by using the dim: 
    ```vb
    CATEGORIES_FILTER
    ```
    -   The script will generate the following lines(please note this is just an example):
    ```vb
    dim CATEGORIES_FILTER
    CATEGORIES_FILTER=CatFilter("CATEGORIES_LIST",lcase(cultureinfo),"",IOM)
    ```
- Update the filters for the questions that need brand filtering by using the dim: 
    ```vb
    BrandFilter
    ```
    -   The script will generate the following lines(please note this is just an example):
    ```vb
    dim BrandFilter

    BrandFilter=CatFilter("BRANDLIST_TEXT_ONLY",lcase(cultureinfo),FLAGCAT.format("a"),IOM)

    dim brand1,brand4,brand2,brand1439,brand10000346,brand3,brand1438,brand10000347,brand180,brand1437,brand1000208,brand1435,ibrand

    for ibrand=0 to IOM.MDM.Types["BRANDLIST_LOGOS"].Elements.Count-1
        execute("brand"+mid(IOM.MDM.Types["BRANDLIST_LOGOS"].Elements[ibrand].Name,1) = "<img src='https://cdn.ipsosinteractive.com/projects/"+IOM.ProjectName+"/logos/"+CText(LCase(CultureInfo))+"/"+mid(IOM.MDM.Types["BRANDLIST_LOGOS"].Elements[ibrand].Name,1)+".jpg' />")
    next
    ```
-   Update the filters for the questions that need statement filtering by using the dim:
    ```vb
    StateFilt_QuestionName
    ```
    -   The script will generate the following lines (please note this is just an example):
    ```vb
    dim StateFilt_BIA,StateFilt_ME,StateFilt_BARCON
    StateFilt_BIA=CatFilter("STATEMENT_LIST_BIA",lcase(cultureinfo),FLAGCAT.format("a"),IOM)
    StateFilt_ME=CatFilter("STATEMENT_LIST_ME",lcase(cultureinfo),FLAGCAT.format("a"),IOM)
    StateFilt_BARCON=CatFilter("STATEMENT_LIST_BARCON",lcase(cultureinfo),FLAGCAT.format("a"),IOM)
    ```
-  Update the filters for the questions that need touchpoint filtering by using the dim
    ```vb
    TouchFilt_QuestionName
    ```
    -   The script will generate the following lines (please note this is just an example):
    ```vb
    dim TouchFilt_RECNONVID,TouchFilt_RECVID
    TouchFilt_RECNONVID=CatFilter("TOUCHPOINT_LIST",lcase(cultureinfo),FLAGCAT.format("a"),IOM)
    TouchFilt_RECVID=CatFilter("TOUCHPOINT_LIST",lcase(cultureinfo),FLAGCAT.format("a"),IOM)
    ```    