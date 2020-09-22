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
- The questions that contain Statement Lists and TouchPoint lists must contain the lists in the following format:<br/>
    STATEMENT_LIST_<span style="background-color: blue">QUESTIONNAME</span><br/>
    TOUCHPOINT_LIST_<span style="background-color: blue">QUESTIONNAME</span>

    * Ex: <br/>
        STATEMENT_LIST_<span style="background-color: blue">BIA</span><br/>
        TOUCHPOINT_LIST_<span style="background-color: blue">RECNONVID</span>
    <br/>    
    The Statement Lists and Touchpoint lists are separated by question
---
## ROUTING
Filters are included in the routing separated by country and category.

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
---
## Languages
-   If there are more language in XML than in MDD, an error is displayed
-   If there is additional language in MDD (not included in XML) that is not en-GB, an error is displayed
-   If there is additional language in MDD (not included in XML) that is en-GB, the script will continue and labels will be included in the predefined lists for en-GB.
---
## Wave variable
> **_NOTE:_**  <font color="red" ><b>Not implemented yet</b></font>
```xml
<wave>
    <name>The label given by the user in XTrack</name>
    <identifier>jn4ch-i385k-lom8-14hdn</identifier>
    <value>8</value>
    <position>4</position>
</wave>

```
- name: The Xtrack wave name will be left up to the researcher to define
- identifier: This is Xtrack’s GUID. Just a value to store in the MDD.  As we move down the road, this will allow us to know what wave data should be associated to.
    > **_NOTE:_**  <font color="cyan" >To be used in naming the delivery zip archive?</font>
- value:&nbsp;&nbsp; <font color="red" ><b>Not used?</b></font>
- position: Not needed for SW/DP
---
## Type of Product variable (for BVC Express)
> **_NOTE:_**  <font color="red" ><b>Not implemented yet</b></font>

Proposed xml:
```xml
<type>1</type>
```
Based on the above value we will set a standard Type question in the routing (in case the tag is part of the xml)
