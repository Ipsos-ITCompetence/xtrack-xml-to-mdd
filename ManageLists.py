from win32com.client import Dispatch
from xml.dom import minidom
import xml.etree.ElementTree as ET
import re
import pandas as pd
import sys
import html
import getch
import numpy
# import os
from glob import glob

#xmlName = 'Multiple Languages\\Wave-BVC Express Wave-ScriptingExportV2.xml'
#mddName ='Multiple Languages\\S19022784.mdd'
xmlName = glob("*.xml")[0]
mddName = glob("*.mdd")[0]

#xmlLower = []
#with open(xmlName, 'r') as f:
#    xmlLower = f.read()

#xmlLower = xmlLower.lower()
#f.closed

#Parse XML
xmltree = ET.parse(xmlName)
root = xmltree.getroot()
#root = ET.fromstring(xmlLower)

def replaceChar(val):
    val = val.replace("’","&#39;")
    val = val.replace("'","&#39;")
    val = val.replace("/","&#47;")

    return val

def LangWorkaround(elem):
    elem = elem.lower()
    if elem in ["en-sa","en-ae","en-eg","en-qa","en-lb","en-jo","en-kw","en-iq","en-pk"]:    	
        elem = "en-cb"
    elif elem in ["fr-ma", "fr-dz"]:
        elem = "fr-mc"
    elif elem == "fil-ph":
        elem = "fil"
    elif elem == "zh-my":
        elem = "zh-mo"	      
    elif elem == "en-ng":
        elem = "en-bz"
    elif elem == "en-hk":  
        elem = "en-zw"       
    elif elem == "fr-re":  
        elem = "br-fr"
    elif elem == "bn-in": 
        elem = "bn"	
    elif elem == "ku-iq":  
        elem = "moh"
    elif elem == "az-az":
        elem = "aze"
    elif elem == "ce-ph":
        elem = "sms"
    elif elem == "ru-lv":
        elem = "se-fi"
    elif elem == "en-ke":
        elem = "co-fr"

    return elem    

for el in root.iter():
    el.tag = el.tag.lower()
    if el.tag == 'scriptlabel':
        if el.text != None and el.text != "_BRANDLIST":
            el.text = el.text.lower()
    for st in el.attrib:
        st = st.lower()
        if st == 'language':
            el.attrib[st] = el.attrib[st].lower()
            el.attrib[st] = LangWorkaround(el.attrib[st])
        if st == 'countryname':
            el.attrib[st] = el.attrib[st].lower()    

def checkSubElements(node, elementsString, multipleInst):
    for el in elementsString.split(","):
        if multipleInst:
            if not (len(node.findall(el,namespaces=None)) > 0):
                print(node.tag + " should have at least 1 " + el + " child node")       
        else:
            if not (len(node.findall(el,namespaces=None)) == 1):
                print(node.tag + " should have exactly 1 " + el + " child node")                   

def checkInnerChild(node, elementsString, multipleInst):
    if node != None:
            checkSubElements(node, elementsString, multipleInst) 

#Check XML Tree
checkSubElements(root, "wave,qtype,categories,countries,lists", False)

#Wave Node
wave = root.find("wave")
if wave != None:
    checkSubElements(wave, "name,identifier,value", False)

#Categories Node
CodesList = []
catsList = root.find("categories")
if catsList != None:
    checkSubElements(catsList, "category", True)          
for item in catsList.findall("category"):
    checkSubElements(item, "labels,description,code", False)
    label = item.find("labels")
    checkInnerChild(label, "label", True)   
    descrip = item.find("description")
    checkInnerChild(descrip, "label", True)
    code = item.find("code")
    if code != None:
        CodesList.append(code.text)
if len(CodesList) != len(set(CodesList)):
    print("Category list has duplicated codes")


#Countries Node
countriesList = root.find("countries")
checkInnerChild(countriesList, "country", True)
for item in countriesList.findall("country"):
    checkSubElements(item, "languages", False)
    lang = item.find("languages")
    checkInnerChild(lang, "language", True)
    if lang != None:
        if len(lang.findall("language",namespaces=None)) > 0:
            for el in lang.findall("language",namespaces=None):
                el.text = LangWorkaround(el.text)

#Lists Node
listsList = root.find("lists")
checkInnerChild(listsList, "list", True) 
for item in listsList.findall("list"):
    checkSubElements(item, "label,scriptlabel,type,questions,listitems", False) #programmingnote
    q = item.find("questions")
    checkInnerChild(q, "question", True)
    li = item.find("listitems")
    checkInnerChild(li, "listitem", True)
    if li != None:
        CodesList.clear()
        for it in li.findall("listitem"):
            checkSubElements(it, "labels,code", False) #mapping
            #attributes
            label = it.find("labels")
            checkInnerChild(label, "label", True)
            #maps = it.find("mapping")   
            #checkInnerChild(maps, "map", True)
            attrib = it.find("attributes")
            checkInnerChild(attrib, "attribute", True)
            code = it.find("code")
            if code != None:
                CodesList.append(code.text)
        if len(CodesList) != len(set(CodesList)):
            print(item.find("scriptlabel").text + " has duplicated codes")

def checkAttrib(node, attributes, root):
    for item in root.iter(node):
        if item.text == None:
            for att in attributes.split(","):
                if att not in item.attrib:
                    print(item.tag + " should have " + att + " attribute")

checkAttrib("label", "language,text", root) 
checkAttrib("question", "label", root) 
checkAttrib("map", "countrycode,countryname,categorycode,categoryname", root) 
checkAttrib("attribute", "label,country,value", root) 

#Checking data format for scripting (against RegEx)
def validFormat(nodes, root):
    for nd in nodes.split(","):
        for item in root.iter(nd):
            if item.text != None:
                result = re.match("(?i)^([a-zA-Z_]{1}\w*)$", item.text)
                if result == None:
                    print(item.tag + " has invalid format: " + item.text)   

validFormat("code,scriptlabel", root)

# #==================================================================================================================
# #============================================    WRITE TO MDD    ==================================================
# #==================================================================================================================
def pause():
  print("Please read the above messages and press any key to continue . . . ")
  getch.getch()

xmlLangs = ""
for country in root.iter("language"):
    xmlLangs = xmlLangs + country.text.lower() + ","
xmlLangs = xmlLangs[:-1]

mdm = Dispatch( 'MDM.Document' )
mdm.Open(mddName, "LATEST", 2)

mddLangs = ""
for language in mdm.Languages:
    mddLangs = mddLangs + language.XMLName.lower() + ","
mddLangs = mddLangs[:-1]

xmlLangArray = xmlLangs.split(",")
mddLangArray = mddLangs.split(",")

errorMsg = ""
XMLonlyLangs = [l for l in xmlLangArray if not (l in mddLangArray) and l != "en-gb"]
if len(XMLonlyLangs) > 0:
    errorMsg = errorMsg + "Language " + str(XMLonlyLangs) + " not present in MDD but present in XML.\n"

ENG_Default = False
MDDonlyLangs = [l for l in mddLangArray if not (l in xmlLangArray) and l != "en-gb"]
if len(MDDonlyLangs) > 0:
    print("Language " + str(MDDonlyLangs) + " not present in XML but present in MDD.\n")
    pause()
    print("ManageLists.py is running\n")  
if "en-gb" in mddLangArray and "en-gb" not in xmlLangArray:
    ENG_Default = True

wLogFile=open('ErrorLog.csv', 'a+', encoding='utf-16')
if errorMsg != "":
    wLogFile.write(errorMsg)
    wLogFile.flush()
    sys.exit("Check ErrorLog.csv for details")

def CreateAdd(lstName, node, currItem):
    mdmObject = [obj for obj in mdm.Types if obj.Name.lower() == lstName.lower()]
    if len(mdmObject) > 0:
        mdm.Types.Remove(mdmObject[0].Name)

    lst = mdm.CreateElements(lstName)
    
    for item in currItem.findall(node):
        catCode = item.find("code").text
        elem = mdm.CreateElement(catCode, "")
        lst.Add(elem) 
	
    mdm.Types.Add(lst)

def SetCatTranslations(findExpression, language, cat, listName, ENG_Default, alternativeLang, ifBrandListName):
    if ifBrandListName != "":
        label = root.findall(findExpression+"[code='"+cat.Name+"']/labels/label[@language='"+language.XMLName.lower()+"']")
        if len(label) == 0:
            label = root.findall(findExpression+"[code='"+cat.Name+"']/labels/label[@language='default']")
            if len(label) == 0 and ENG_Default:
                label = root.findall(findExpression+"[code='"+cat.Name+"']/labels/label[@language='"+alternativeLang+"']") 
        
        labelTxt = replaceChar(html.escape(label[0].attrib['text']))
        if listName == ifBrandListName+"_TEXT_ONLY":
            cat.Labels.Text = labelTxt
        elif listName == ifBrandListName+"_LOGOS_LBT":
            cat.Labels.Text = "{#brand" + cat.Name.replace("_","") + "}<br/><span style='display:none'>" + labelTxt + "</span>"
        elif listName == ifBrandListName+"_LOGOS": 
            cat.Labels.Text = "<center>{#brand" + cat.Name.replace("_","") + "}<br/>" + labelTxt + "</center>"                     
        elif listName == ifBrandListName+"_CLOSENESS":        
            cat.Labels.Text = "{#brand" + cat.Name.replace("_","") + "}<br/><span class='closeness-hidden-content'>" + labelTxt + "</span>"
    else:
        label = root.findall(findExpression + "[@language='"+language.XMLName.lower()+"']")
        if len(label)>0:
            cat.Labels.Text = replaceChar(html.escape(label[0].attrib['text']))
        else:
            label = root.findall(findExpression + "[@language='default']")
            if len(label) == 0 and ENG_Default:
                label = root.findall(findExpression + "[@language='"+alternativeLang+"']") 
            if len(label)>0:
                cat.Labels.Text = replaceChar(html.escape(label[0].attrib['text']))

branddim = ""

#Creating Category List
CreateAdd("CATEGORIES_LIST", "categories/category", root)

#Creating Lists
brandListsArray = []
for list in root.iter("list"):
    listName = list.find("scriptlabel").text
    if list.find("type").text.lower() == "brand":
        brandListsArray.append(listName[1:])
        CreateAdd(listName[1:]+"_TEXT_ONLY", "listitems/listitem", list)
        CreateAdd(listName[1:]+"_LOGOS_LBT", "listitems/listitem", list)
        CreateAdd(listName[1:]+"_LOGOS", "listitems/listitem", list)
        CreateAdd(listName[1:]+"_CLOSENESS", "listitems/listitem", list)    
        for item in list.findall("listitems/listitem"):
            catCode = item.find("code").text
            if (catCode[1:] != "990" and catCode[1:] != "998" and catCode[1:] != "999"):
                branddim=branddim+"brand"+catCode[1:]+","
    else:    
        CreateAdd(listName[1:], "listitems/listitem", list)
       
mdm.Save() 

# #Adding Lists Translations
alternativeLang = ""
if ENG_Default:
    alternativeLang = xmlLangArray[0]
for language in mdm.Languages:
    if not(language.xmlName.lower() in xmlLangArray) and language.xmlName.lower() != "en-gb":
        continue
    mdm.Languages.Current=language

    #Update wave and qtype
    wave = root.find("wave")
    mdm.Items["WAVE_NAME"].Label = wave.find("name").text
    mdm.Items["WAVE_IDENTIFIER"].Label = wave.find("identifier").text
    mdm.Items["Wave"].Label = wave.find("value").text
    #mdm.Items["QTYPE"].Label = root.find("qtype").text    

    ExplInsert = False
    mdmObject = [obj for obj in mdm.Fields if obj.Name == "EXPLANATION_INSERT"]
    if len(mdmObject) > 0:
        ExplInsert = True
    for item in mdm.Types:
        ifBrandListName = item.FullName.replace("_TEXT_ONLY","").replace("_LOGOS_LBT","").replace("_LOGOS","").replace("_CLOSENESS","")
        for cat in item.Elements:
            if item.FullName == "CATEGORIES_LIST":        
                SetCatTranslations("categories/category[code='"+cat.Name+"']/labels/label", language, cat, item.FullName, ENG_Default, alternativeLang, "")      
                if ExplInsert:
                    mdmCat = [c for c in mdmObject[0].Elements if c.Name == cat.Name] 
                    label = root.findall("categories/category[code='"+cat.Name+"']/description/label[@language='"+language.XMLName.lower()+"']")
                    if len(label) == 0:
                        label = root.findall("categories/category[code='"+cat.Name+"']/description/label[@language='default']")
                        if len(label) == 0 and ENG_Default:
                            label = root.findall("categories/category[code='"+cat.Name+"']/description/label[@language='"+alternativeLang+"']") 
        
                    if len(label) > 0:
                        if len(mdmCat) > 0:
                            mdmCat[0].Label = label[0].attrib["text"]
                        else:
                            elem = mdm.CreateElement(cat.Name, label[0].attrib["text"])
                            mdmObject[0].Elements.Add(elem)                                   
            elif ifBrandListName in brandListsArray:
                SetCatTranslations("lists/list[scriptlabel='_" + ifBrandListName + "']/listitems/listitem", language, cat, item.FullName, ENG_Default, alternativeLang, ifBrandListName)                                     
            else:                              
                lstName = root.findall("lists/list[scriptlabel='_"+item.FullName+"']")
                if len(lstName)>0:
                    SetCatTranslations("lists/list[scriptlabel='_"+item.FullName+"']/listitems/listitem[code='"+cat.Name+"']/labels/label", language, cat, item.FullName, ENG_Default, alternativeLang, "")

    if ExplInsert:
        for cat in mdmObject[0].Elements:
            mdmCat = [c for c in mdm.Types["CATEGORIES_LIST"].Elements if c.Name == cat.Name]
            if len(mdmCat) == 0:
                mdmObject[0].Elements.Remove(cat.Name)    

    mdm.Save()
    print("Translations for " + language.XMLName + " are added!")

#Update routing
writeFilter = ""
writeFilter = writeFilter + "QTYPE.Response.Value={_" + root.find("qtype").text  + "}\n"
writeFilter = writeFilter + "WAVE_NAME.Response.Value=\"" + wave.find("name").text + "\"\n"
writeFilter = writeFilter + "WAVE_IDENTIFIER.Response.Value=\"" + wave.find("identifier").text + "\"\n"
writeFilter = writeFilter + "Wave.Response.Value=\""+wave.find("value").text + "\"\n"
if root.find("flagtimescale") != None:
    writeFilter = writeFilter + "FLAGTIMESCALE.Response.Value={_"+root.find("flagtimescale").text + "}\n"
if root.find("flagprovider") != None:
    writeFilter = writeFilter + "FLAGPROVIDER.Response.Value={_"+root.find("flagprovider").text + "}\n"
writeFilter = writeFilter + "\n"

branddim=branddim+"ibrand"
writeFilter = writeFilter + "dim " + branddim + "\n\n"
writeFilter = writeFilter + "For ibrand=0 to IOM.MDM.Types[\"BRANDLIST_LOGOS\"].Elements.Count-1\n\t"
writeFilter = writeFilter + "execute(\"brand\"+mid(IOM.MDM.Types[\"BRANDLIST_LOGOS\"].Elements[ibrand].Name,1) +\" = \"\"<img src='https://cdn.ipsosinteractive.com/projects/\"+IOM.ProjectName+\"/img/\"+CText(LCase(CultureInfo))+\"/logos/\"+mid(IOM.MDM.Types[\"BRANDLIST_LOGOS\"].Elements[ibrand].Name,1)+\".jpg' />\"\"\")\n"
writeFilter = writeFilter + "Next\n\n"

writeFilter = writeFilter + "Select Case COUNTRY_\n"
dims = "Dim "
for cc in root.iter("country"):
    writeFilter = writeFilter + "\tCase {" + cc.attrib["code"] + "}\n"
    # listMapping[cc.attrib["code"]] = {}
    for cat in root.iter("category"):
        writeFilter = writeFilter + "\t\tif FLAGCAT.ContainsAny({" + cat.find("code").text + "}) Then\n"
        # listMapping[cc.attrib["code"]][cat.find("code").text] = {}
        for lst in root.iter("list"):
            filterVar = lst.find("scriptlabel").text[1:] #+ "_Filter"
            if dims.find(filterVar + ",") == -1:
                dims = dims + filterVar + ","
            # listMapping[cc.attrib["code"]][cat.find("code").text][filterVar] = ""
            val = ""
            for it in lst.findall("listitems/listitem"):           
                map = it.findall("mapping/map[@countrycode=\"" + cc.attrib["code"] + "\"][@categorycode=\"" + cat.find("code").text + "\"]")
                if len(map) > 0:
                    val = val + it.find("code").text + ","

            # listMapping[cc.attrib["code"]][cat.find("code").text][filterVar] = val[:-1]
            if val != "":
                writeFilter = writeFilter + "\t\t\t" + filterVar + "=" + filterVar + " + {" + val[:-1] + "}\n"

        writeFilter = writeFilter + "\t\tEnd If\n"    
writeFilter = writeFilter + "End Select\n"

#Add custom filters from attributes
attribFilter = {}
attribDims = ""
for lst in root.iter("list"):
    for it in lst.findall("listitems/listitem"):
        for att in it.findall("attributes/attribute"):
            countryName = att.attrib['countrycode'].lower()
            if not (countryName in attribFilter):
                attribFilter[countryName] = {}    
            fltrName = (att.attrib['label'] + "_" + att.attrib['value']).lower()   
            if not (fltrName in attribFilter[countryName]):
                attribFilter[countryName][fltrName] = it.find("code").text
                if not ("," + fltrName in attribDims):
                    attribDims = attribDims + "," + fltrName
            else:
                attribFilter[countryName][fltrName] += "," + it.find("code").text 

writeFilter = writeFilter + "Dim " + attribDims[1:] + "\n"
writeFilter = writeFilter + "Select Case COUNTRY_\n"
for key in attribFilter:
    if key != 'default':
        #writeFilter = writeFilter + "\tCase {" + root.find("countries/country[@countryname='"+key+"']").attrib['code'] + "}\n"
        writeFilter = writeFilter + "\tCase {" + key + "}\n"
        for fltr in attribFilter[key]:
            writeFilter = writeFilter + "\t\t" + fltr + " = {" + attribFilter[key][fltr] + "}\n"
if 'default' in attribFilter:
    writeFilter = writeFilter + "\tCase Else\n"
    for fltr in attribFilter['default']:
        writeFilter = writeFilter + "\t\t" + fltr + " = {" + attribFilter['default'][fltr] + "}\n"  
writeFilter = writeFilter + "End Select\n"

#Mapping-group filters
TPfilter_arr = {}
for cc in root.iter("mapping-groups"):
    for mg in cc.findall("mapping-group/mappings"):
        tpFilter = mg.find("scriptlabel").attrib["tp_label"].lower()[1:]
        for tp in root.findall("lists/list[scriptlabel='_"+tpFilter+"']/listitems/listitem/code"):
            if not tpFilter in TPfilter_arr:
                TPfilter_arr[tpFilter] = {}
            filterCats = ','.join([str(elem) for elem in [st.attrib['st_code'] for st in cc.findall("mapping-group/mappings[@tp_code='" + tp.text + "']/scriptlabel[@tp_label='"+mg.find("scriptlabel").attrib["tp_label"]+"']..mapping")]])     
            if filterCats != "":
                if not tp.text in TPfilter_arr[tpFilter]:
                    TPfilter_arr[tpFilter][tp.text] = {}                
                TPfilter_arr[tpFilter][tp.text][mg.find("scriptlabel").attrib["st_label"].lower()[1:]] = "{" + filterCats + "}"
            
for filter in TPfilter_arr:
    writeFilter = writeFilter + "Select Case " + filter + "\n"
    for cat in TPfilter_arr[filter]:
        writeFilter = writeFilter + "\tCase {" + cat + "}\n"
        for stlist in TPfilter_arr[filter][cat]:
            if dims.find(stlist+",") != -1: 
                writeFilter = writeFilter + "\t\t" + stlist + " = " + stlist + "*" + TPfilter_arr[filter][cat][stlist] + "\n"
            else:
                dims = dims + stlist + "," 
                writeFilter = writeFilter + "\t\t" + stlist + " = " + TPfilter_arr[filter][cat][stlist] + "\n"    
    writeFilter = writeFilter + "End Select\n"

routingscript = mdm.Routing.Script

if  routingscript.find("'*** Start--List--Filters ***")!=-1 and \
        routingscript.find("'*** End--List--Filters ***")!=-1:

    routingscript = routingscript.replace(routingscript[routingscript.find("'*** Start--List--Filters ***"):routingscript.find("'*** End--List--Filters ***") + len("'*** End--List--Filters ***")], "'*** Start--List--Filters ***\n" + dims[:-1] + "\n" + writeFilter + "'*** End--List--Filters ***")
mdm.Routing.Script = routingscript

mdm.Save()
mdm.Close() 