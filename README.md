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
    Select Case CountryCode
        Case "FR"
            CATEGORIES_FILTER = {_1,_2}
        Case "ES"
            CATEGORIES_FILTER = {_1,_2}
    End Select
    ```
- Update the filters for the questions that need brand filtering by using the dim: 
    ```vb
    BrandFilter
    ```
    -   The script will generate the following lines(please note this is just an example):
    ```vb
    dim BrandFilter
    Select Case CountryCode
        Case "FR"
            BrandFilter = {}
            if FLAGCAT.ContainsAny({_1}) Then BrandFilter = BrandFilter + {_4,_5,_6,_8,_14,_17,_18,_28,_31,_32}
            if FLAGCAT.ContainsAny({_2}) Then BrandFilter = BrandFilter + {_2,_7,_22,_23,_25,_26,_27,_29,_30}
        Case "ES"
            BrandFilter = {}
            if FLAGCAT.ContainsAny({_1}) Then BrandFilter = BrandFilter + {_10,_12,_13,_15,_16,_20,_24}
            if FLAGCAT.ContainsAny({_2}) Then BrandFilter = BrandFilter + {_1,_9,_11,_13,_16,_19,_21,_22,_27,_30}
    End Select

    dim brand2,brand4,brand5,brand6,brand7,brand8,brand14,brand17,brand18,brand22,brand23,brand25,brand26,brand27,brand28,brand29,brand30,brand31,brand32,brand1,brand9,brand10,brand11,brand12,brand13,brand15,brand16,brand19,brand20,brand21,brand24,ibrand

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
    dim StateFilt_ME
    Select Case CountryCode
        Case "FR"
            StateFilt_ME = {}
            if FLAGCAT.ContainsAny({_2}) Then StateFilt_ME = StateFilt_ME + {_62,_63,_64,_65,_66,_67,_68,_69,_70,_71,_72,_73,_74,_75}
            if FLAGCAT.ContainsAny({_1}) Then StateFilt_ME = StateFilt_ME + {_62,_63,_64,_65,_66,_67,_68,_69,_70,_71,_72,_73,_74,_75}
        Case "ES"
            StateFilt_ME = {}
            if FLAGCAT.ContainsAny({_2}) Then StateFilt_ME = StateFilt_ME + {_62,_63,_64,_65,_66,_67,_68,_69,_70,_71,_72,_73,_74,_75}
            if FLAGCAT.ContainsAny({_1}) Then StateFilt_ME = StateFilt_ME + {_62,_63,_64,_65,_66,_67,_68,_69,_70,_71,_72,_73,_74,_75}
    End Select
    ```
-  Update the filters for the questions that need touchpoint filtering by using the dim
    ```vb
    TouchFilt_QuestionName
    ```
    -   The script will generate the following lines (please note this is just an example):
    ```vb
    dim TouchFilt_RECNONVID
    Select Case CountryCode
        Case "FR"
            TouchFilt_RECNONVID = {}
            if FLAGCAT.ContainsAny({_1}) Then TouchFilt_RECNONVID = TouchFilt_RECNONVID + {_1,_2,_3}
    End Select
    ```    
---
## Languages
-   If there are more language in XML than in MDD, an error is displayed
-   If there is additional language in MDD (not included in XML) that is not en-GB, an error is displayed
-   If there is additional language in MDD (not included in XML) that is en-GB, the script will continue and labels will be included in the predefined lists for en-GB.
---
## Wave variable
```xml
<wave>
    <name>The label given by the user in XTrack</name>
    <identifier>jn4ch-i385k-lom8-14hdn</identifier>
    <value>8</value>
</wave>

```
- name: The Xtrack wave name will be left up to the researcher to define
- identifier: This is Xtrack’s GUID. Just a value to store in the MDD.  As we move down the road, this will allow us to know what wave data should be associated to.
    > **_NOTE:_**  <font color="cyan" >To be used in naming the delivery zip archive?</font>
- value:&nbsp;&nbsp; <font color="red" ><b>Not used?</b></font>
- position: Not needed for SW/DP
---
## Type of Product variable (for BVC Express)

Proposed xml:
```xml
<qtype>1</qtype>
```
Based on the above value we will set a standard Type question in the routing (in case the tag is part of the xml)
