#target InDesign
/*
Release notes: 2.1 - added in paragraph sequencing function, goes through document and sorts objects by page and then geometric bounds, this is to ensure that media elements are properly sequenced with their paragraph	
	*/

loadPrototypeFunctions();
/* set global variables */
	//error report
	var errorReport = [];
	//filelist
	var myFiles = [];
	//missingstyles report
	var missingStyles = [];

/* initialize namespaces */
var xlink = new Namespace("http://www.w3.org/1999/xlink");
var xmlns = new Namespace("http://www.w3.org/1999/xhtml");

/* load document metadata */
var metaArr = {};
var fileObj = File("C:\\Users\\kstaples\\Desktop\\0001 - Projects\\23 - NLM Conversion\\Metadata3rdto5thClass.csv");
if(!fileObj.exists){
    alert("Metadata file does not exist.");
    }
var delimiter = new RegExp(',');
fileObj.open('r');
while ( ! fileObj.eof){
    var currentLine = fileObj.readln();
	var arrMetadata = splitCSVButIgnoreCommasInDoublequotes (currentLine) // http://stackoverflow.com/questions/11456850/split-a-string-by-commas-but-ignore-commas-within-double-quotes-using-javascript
	metaArr[arrMetadata[0]] = {sectionTitle:arrMetadata[1],bookid:arrMetadata[2], booktitle:arrMetadata[3], parttitle:arrMetadata[4]};
    }
if(typeof iFigure === 'undefined'){manifest = []}
getDocumentList ();
bContinue = true; // boolean for skipping documents for testing
for(var i = 0; i<myFiles.length; i++){
	if(File(myFiles[i]).name.indexOf("U.indd")!=-1||File(myFiles[i]).name.indexOf("TOC")!=-1||File(myFiles[i]).name.indexOf("_U")!=-1||File(myFiles[i]).name.indexOf("CEPE")==-1||File(myFiles[i]).name.indexOf("_A")!=-1){
		continue;
		}
	var oldInteractionPrefs = app.scriptPreferences.userInteractionLevel;
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
	 
	 // for targetting specific files / ranges	 
	 /*
	if(File(myFiles[i]).name=="CEPE05200B1_010.indd"){
		bContinue = false;
		}
	if(bContinue == true){
		continue;
		}
		*/
		
	app.open(File(myFiles[i]),true);
	convertBulletsAndNumberingToText ();
	chapterFolder = "C:\\Users\\kstaples\\Desktop\\0001 - Projects\\47 - New NLM\\NLM Out\\"+app.activeDocument.name.replace(".indd","")+"\\";
	Folder(chapterFolder).create();
	var metaID = app.activeDocument.name.replace(".indd","");
	if(!metaArr[metaID]){ // check if current document has existing metadata object
		errorReport[errorReport.length] = app.activeDocument.name+", does not have metadata.";
		app.activeDocument.close(SaveOptions.NO);
		app.scriptPreferences.userInteractionLevel = oldInteractionPrefs;
		continue;
		}
	else
	{
		var doc = app.activeDocument;
		convertMathToolsV1DocumentToMathToolsV2 (doc);
		getChapterVariables ();
		initializeXmlBook();
		parseParagraph();
		}
    var xmlFile = File(chapterFolder+app.activeDocument.name.replace(".indd","")+".xml");
    xmlFile.open('w');
    xmlFile.encoding = 'UTF-8';
    xmlFile.write("""<?xml version="1.0" encoding="UTF-8"?>"""+"\r"+"<!DOCTYPE book SYSTEM \"d:\\parser\\nlm-book\\book.dtd\">"+"\r"+book.toXMLString().replace(/&lt;/gi,"<").replace(/&gt;/gi,">").replace(/inlinegraphic/gi,"inline-graphic").replace(/mimesubtype/gi,"mime-subtype").replace(/§\/bold§/gi,"</bold>").replace(/§bold§/gi,"<bold>").replace(/xmlnsxlink/gi,"xmlns:xlink").replace(/xlinkhref/gi,"xlink:href").replace(/xlinktitle/gi,"xlink:title").replace(/xlinkhref/gi,"xlink:href").replace(/listitem/gi,"list-item").replace(/xlinkmime-subtype/gi,"xlink:mime-subtype"));
    xmlFile.close();
	createManifest(manifest);
	app.activeDocument.close(SaveOptions.NO);
	app.scriptPreferences.userInteractionLevel = oldInteractionPrefs;
	}

function initializeXmlBook(){
	resetGlobals();
	$.writeln("entering: initializeXmlBook");
	iVol = 0;
	iPart = 0;
	iSection = 0;
	book  = new XML("<book>");
	book.appendChild(new XML("<book-meta>"));
	//book.appendChild(new XML("<book-front>"));
	book['book-meta'].appendChild(new XML("<book-id>"));
	book['book-meta']['book-id'] = metaArr[metaID].bookid;
	book['book-meta'].appendChild(new XML("<book-title-group>"));
	book['book-meta']['book-title-group']['book-title'] = metaArr[metaID].booktitle;
	book.appendChild(new XML("<body>"));
	//setup book parts
	//create volume & set attributes
	iVol++;
	book.body.appendChild(new XML("<book-part>"));
	var volume = book.body['book-part'];
	volume.@id = "VOL"+pad(iVol,3);
	volume.@['book-part-type']="volume";
    volume.@['book-part-number']=1;
    volume.appendChild(new XML("<book-part-meta>"));
    volume['book-part-meta'].appendChild(new XML("<title-group>"));
    volume['book-part-meta']['title-group'].title = "Volume "+iVol;
    volume.appendChild(new XML("<body>"));
    //create part & set attributes
    iPart++;
    volume.body.appendChild(new XML("<book-part>"));
    var part = volume.body['book-part'];
    part.@id = "PART"+pad(iPart,3);
    part.@['book-part-type']="part";
    part.@['book-part-number']=1;
    part.@['xlink-role']="volume"+iVol;
    part.appendChild(new XML("<book-part-meta>"));
    part['book-part-meta'].appendChild(new XML("<title-group>"));
    part['book-part-meta']['title-group'].title = metaArr[metaID].parttitle;
    part.appendChild(new XML("<body>"));
    //create section & set attributes
    iSection++;
    part.body.appendChild(new XML("<book-part>"));
    var section = part.body['book-part'];
    section.@id = "SEC"+pad(iSection,3);
    section.@['book-part-type'] = "section";
    section.@['book-part-number'] = iChapter;
    section.appendChild(new XML("<book-part-meta>"));
    section['book-part-meta'].appendChild(new XML("<title-group>"));
    section.@['xlink-role'] = "part"+iPart;
    section['book-part-meta']['title-group'].title = metaArr[metaID].sectionTitle;
    section.appendChild(new XML("<body>"));
    //create chapter & set attributes
    section.body.appendChild(new XML("<book-part>"));
    chapter = section.body['book-part'];
    if(iChapter.length == 3){
    chapter.@id = "CH"+pad(iChapter,4);
    }else{chapter.@id = "CH"+pad(iChapter,4)}
    chapter.@['book-part-type'] = "chapter"; //+g.iChapter;
    chapter.@['book-part-number'] = 1;
    chapter.appendChild(new XML("<book-part-meta>"));
    chapter['book-part-meta'].appendChild(new XML("<title-group>"));
    chapter.@['xlink-role']="chapter"+iChapter; //+g.iSection;
    chapter['book-part-meta']['title-group'].title.italic = "Chapter "+iChapter+" "+strChapterTitle;
    chapter.appendChild(new XML("<body>"));
    chapter = chapter.body;
	}

function getDocumentList(){
	$.writeln("entering: getDocumentList ");
//var folders = ["C:\\Users\\kstaples\\Desktop\\Edition 3\\InDesign","C:\\Users\\kstaples\\Desktop\\0001 - Projects\\23 - NLM Conversion\\1-CEPE03200_CURRENT MASTER FILES_activated  Feb10-15","C:\\Users\\kstaples\\Desktop\\0001 - Projects\\23 - NLM Conversion\\4th Class Corrected Formulas"];
//var folder = Folder.selectDialog ();
//var folders = ["C:\\Users\\kstaples\\Desktop\\Edition 3\\InDesign"];
var folders = ["C:\\Users\\kstaples\\Desktop\\0001 - Projects\\23 - NLM Conversion\\1-CEPE03200_CURRENT MASTER FILES_activated  Feb10-15"];
//var folder = Folder("C:\\Users\\kstaples\\Desktop\\0001 - Projects\\23 - NLM Conversion\\Isolated Test");
//var folder = Folder("C:\\Users\\kstaples\\Desktop\\0001 - Projects\\23 - NLM Conversion\\4th Class Corrected Formulas")

for(var i = 0; i<folders.length; i++){
    var folder = Folder(folders[i]);
    GetSubFolders(folder);
    }
	
/*
myFiles = ["C:\\Users\\kstaples\\Desktop\\Edition 3\\InDesign\\Book 2-Edition 3\\Unit 5 - Chapter 23-26\\CEPE05200B2_023.indd"]; // over write myFile list with single document
myFiles = ["C:\\Users\\kstaples\\Desktop\\Edition 3\\InDesign\\Book 2-Edition 3\\Unit 6 - Chapter 27-31\\CEPE05200B2_027.indd"];
*/
}

function convertMathToolsV1DocumentToMathToolsV2(doc){
	$.writeln("entering: convertMathToolsV1DocumentToMathToolsV2");
    //enable math for document
    app.activeDocument.mtEnableMathToolsOnDocument();
    //convert document to mathtools v2
    app.activeDocument.mtConvertDocToMathToolsV2();
    }

function loadPrototypeFunctions(){ // anonymous function extends Array functions to include indexOf
	if (!Array.prototype.indexOf) {
		Array.prototype.indexOf = function(searchElement, fromIndex) {
			var k;
			// 1. Let o be the result of calling ToObject passing
			//    the this value as the argument.
			if (this == null) {
				throw new TypeError('"this" is null or not defined');
				}
			var o = Object(this);
			// 2. Let lenValue be the result of calling the Get
			//    internal method of o with the argument "length".
			// 3. Let len be ToUint32(lenValue).
			var len = o.length >>> 0;
			// 4. If len is 0, return -1.
			if (len === 0) {
				return -1;
				}
			// 5. If argument fromIndex was passed let n be
			//    ToInteger(fromIndex); else let n be 0.
			var n = +fromIndex || 0;
			if (Math.abs(n) === Infinity) {
				n = 0;
				}
			// 6. If n >= len, return -1.
			if (n >= len) {
				return -1;
				}
			// 7. If n >= 0, then Let k be n.
			// 8. Else, n<0, Let k be len - abs(n).
			//    If k is less than 0, then let k be 0.
			k = Math.max(n >= 0 ? n : len - Math.abs(n), 0);
			// 9. Repeat, while k < len
			while (k < len) {
				// a. Let Pk be ToString(k).
				//   This is implicit for LHS operands of the in operator
				// b. Let kPresent be the result of calling the
				//    HasProperty internal method of o with argument Pk.
				//   This step can be combined with c
				// c. If kPresent is true, then
				//    i.  Let elementK be the result of calling the Get
				//        internal method of o with the argument ToString(k).
				//   ii.  Let same be the result of applying the
				//        Strict Equality Comparison Algorithm to
				//        searchElement and elementK.
				//  iii.  If same is true, return k.
				if (k in o && o[k] === searchElement) {
					return k;
					}
				k++;
				}
			return -1;
			}
		}
	}

function splitCSVButIgnoreCommasInDoublequotes(str) {  
	$.writeln("splitCSVButIgnoreCommasInDoublequotes");
    //split the str first  
    //then merge the elments between two double quotes  
    var delimiter = ',';  
    var quotes = '"';  
    var elements = str.split(delimiter);  
    var newElements = [];  
    for (var i = 0; i < elements.length; ++i) {  
        if (elements[i].indexOf(quotes) >= 0) {//the left double quotes is found  
            var indexOfRightQuotes = -1;  
            var tmp = elements[i];  
            //find the right double quotes  
            for (var j = i + 1; j < elements.length; ++j) {  
                if (elements[j].indexOf(quotes) >= 0) {  
                    indexOfRightQuotes = j;  
                }  
            }  
            //found the right double quotes  
            //merge all the elements between double quotes  
            if (-1 != indexOfRightQuotes) {   
                for (var j = i + 1; j <= indexOfRightQuotes; ++j) {  
                    tmp = tmp + delimiter + elements[j];  
                }  
                newElements.push(tmp);  
                i = indexOfRightQuotes;  
            }  
            else { //right double quotes is not found  
                newElements.push(elements[i]);  
            }  
        }  
        else {//no left double quotes is found  
            newElements.push(elements[i]);  
        }  
    }  
    return newElements;  
} 

function pad(num, size) {
    var s = num+"";
    while (s.length < size) s = "0" + s;
    return s;
    }

function getChapterVariables(){
	$.writeln("entering: getChapterVariables");
    for(var i = 0; i<app.activeDocument.pages[0].textFrames.length; i++){
        var tf = app.activeDocument.pages[0].textFrames[i];
        for(var j = 0; j<tf.paragraphs.length; j++){
            var paragraph = tf.paragraphs[j];
            if(paragraph.appliedParagraphStyle.name=="CHAPTER"||paragraph.appliedParagraphStyle.name=="Chapter #"){
                iChapter = getContents(paragraph).replace(/[^0-9]/gi,"");
                }
            if(paragraph.appliedParagraphStyle.name=="Chapter Title"){
                strChapterTitle = getContents(paragraph);
                }
            }
        }
    }

function getContents(p){
	$.writeln("entering: getContents");
    var paragraphContents = "";
    if(true){
        arrMathZonesMML = [];
        var stateSup = false;
        var stateBold = false;
        var tmpStateBold = false;
        var stateItalic = false;
        var tmpStateBold = false;
        var counterMathZone = 0;
        var mathZones = p.mtAllMathZones; // math zones
        for(var i = 0; i<mathZones.length; i++){
            arrMathZonesMML[arrMathZonesMML.length] = mathZones[i]; // stores all math zones for paragraph
            }
        for(var i = 0; i<p.characters.length; i++){
			var chr = p.characters[i];
			var chrContents = chr.contents; // variable space for character contents
			if(chr.contents.constructor.name === "Enumerator"){
				//$.writeln("Enumerator section entered: "+chr.contents.toString()+", length:"+chr.contents.toString().length);
				chrContents = chr.contents;
				if(chr.contents==SpecialCharacters.ARABIC_COMMA){chrContents = "";}
				if(chr.contents==SpecialCharacters.ARABIC_KASHIDA){chrContents = "";}
				if(chr.contents==SpecialCharacters.ARABIC_QUESTION_MARK){chrContents = "";}
				if(chr.contents==SpecialCharacters.ARABIC_SEMICOLON){chrContents = "";}
				if(chr.contents==SpecialCharacters.AUTO_PAGE_NUMBER){chrContents = "";}
				if(chr.contents==SpecialCharacters.BULLET_CHARACTER){chrContents = "&#8226;";}
				if(chr.contents==SpecialCharacters.COLUMN_BREAK){chrContents = "";}
				if(chr.contents==SpecialCharacters.COPYRIGHT_SYMBOL){chrContents = "";}
				if(chr.contents==SpecialCharacters.DEGREE_SYMBOL){chrContents = "";}
				if(chr.contents==SpecialCharacters.DISCRETIONARY_HYPHEN){chrContents = "";}
				if(chr.contents==SpecialCharacters.DISCRETIONARY_LINE_BREAK){chrContents = "";}
				if(chr.contents==SpecialCharacters.DOUBLE_LEFT_QUOTE){chrContents = "&ldquo;";}
				if(chr.contents==SpecialCharacters.DOUBLE_RIGHT_QUOTE){chrContents = "&rdquo;";}
				if(chr.contents==SpecialCharacters.DOUBLE_STRAIGHT_QUOTE){chrContents = "&quot;";}
				if(chr.contents==SpecialCharacters.ELLIPSIS_CHARACTER){chrContents = "&hellip;";}
				if(chr.contents==SpecialCharacters.EM_DASH){chrContents = "";}
				if(chr.contents==SpecialCharacters.EM_SPACE){chrContents = "";}
				if(chr.contents==SpecialCharacters.END_NESTED_STYLE){chrContents = "";}
				if(chr.contents==SpecialCharacters.EN_DASH){chrContents = "";}
				if(chr.contents==SpecialCharacters.EN_SPACE){chrContents = "";}
				if(chr.contents==SpecialCharacters.EVEN_PAGE_BREAK){chrContents = "";}
				if(chr.contents==SpecialCharacters.FIGURE_SPACE){chrContents = "";}
				if(chr.contents==SpecialCharacters.FIXED_WIDTH_NONBREAKING_SPACE){chrContents = "";}
				if(chr.contents==SpecialCharacters.FLUSH_SPACE){chrContents = "";}
				if(chr.contents==SpecialCharacters.FOOTNOTE_SYMBOL){chrContents = "";}
				if(chr.contents==SpecialCharacters.FORCED_LINE_BREAK){chrContents = "";}
				if(chr.contents==SpecialCharacters.FRAME_BREAK){chrContents = "";}
				if(chr.contents==SpecialCharacters.HAIR_SPACE){chrContents = "";}
				if(chr.contents==SpecialCharacters.HEBREW_GERESH){chrContents = "";}
				if(chr.contents==SpecialCharacters.HEBREW_GERSHAYIM){chrContents = "";}
				if(chr.contents==SpecialCharacters.HEBREW_MAQAF){chrContents = "";}
				if(chr.contents==SpecialCharacters.INDENT_HERE_TAB){chrContents = "";}
				if(chr.contents==SpecialCharacters.LEFT_TO_RIGHT_MARK){chrContents = "";}
				if(chr.contents==SpecialCharacters.NEXT_PAGE_NUMBER){chrContents = "";}
				if(chr.contents==SpecialCharacters.NONBREAKING_HYPHEN){chrContents = "";}
				if(chr.contents==SpecialCharacters.NONBREAKING_SPACE){chrContents = "";}
				if(chr.contents==SpecialCharacters.ODD_PAGE_BREAK){chrContents = "";}
				if(chr.contents==SpecialCharacters.PAGE_BREAK){chrContents = "";}
				if(chr.contents==SpecialCharacters.PARAGRAPH_SYMBOL){chrContents = "";}
				if(chr.contents==SpecialCharacters.PREVIOUS_PAGE_NUMBER){chrContents = "";}
				if(chr.contents==SpecialCharacters.PUNCTUATION_SPACE){chrContents = "";}
				if(chr.contents==SpecialCharacters.QUARTER_SPACE){chrContents = "";}
				if(chr.contents==SpecialCharacters.REGISTERED_TRADEMARK){chrContents = "";}
				if(chr.contents==SpecialCharacters.RIGHT_INDENT_TAB){chrContents = "";}
				if(chr.contents==SpecialCharacters.RIGHT_TO_LEFT_MARK){chrContents = "";}
				if(chr.contents==SpecialCharacters.SECTION_MARKER){chrContents = "";}
				if(chr.contents==SpecialCharacters.SECTION_SYMBOL){chrContents = "";}
				if(chr.contents==SpecialCharacters.SINGLE_LEFT_QUOTE){chrContents = "";}
				if(chr.contents==SpecialCharacters.SINGLE_RIGHT_QUOTE){chrContents = "";}
				if(chr.contents==SpecialCharacters.SINGLE_STRAIGHT_QUOTE){chrContents = "";}
				if(chr.contents==SpecialCharacters.SIXTH_SPACE){chrContents = "";}
				if(chr.contents==SpecialCharacters.TEXT_VARIABLE){chrContents = "";}
				if(chr.contents==SpecialCharacters.THIN_SPACE){chrContents = "";}
				if(chr.contents==SpecialCharacters.THIRD_SPACE){chrContents = "";}
				if(chr.contents==SpecialCharacters.TRADEMARK_SYMBOL){chrContents = "";}
				if(chr.contents==SpecialCharacters.ZERO_WIDTH_NONJOINER){chrContents = "";}
				}
            var insertionPoint = p.insertionPoints[i]; // target insertion point
            // 0 not math, 1 mathzone start, 2 end marker, 3 index inside mathzone
            // is character part of math zone
            //start of math zone
            if(insertionPoint.mtIsMath()==1){
				if(mathZones[counterMathZone].mtExportMathZoneAsMathML().substring(mathZones[counterMathZone].mtExportMathZoneAsMathML().indexOf("<mat"),mathZones[counterMathZone].mtExportMathZoneAsMathML().length).indexOf("</math>")==-1){
					errorReport[errorReport.length] = app.activeDocument.name+", "+p.parentTextFrames[0].parentPage.name;
					paragraphContents += mathZones[counterMathZone].mtExportMathZoneAsMathML().substring(mathZones[counterMathZone].mtExportMathZoneAsMathML().indexOf("<mat"),mathZones[counterMathZone].mtExportMathZoneAsMathML().length)+"</math>";
					}
				else
				{
					paragraphContents += mathZones[counterMathZone].mtExportMathZoneAsMathML().substring(mathZones[counterMathZone].mtExportMathZoneAsMathML().indexOf("<mat"),mathZones[counterMathZone].mtExportMathZoneAsMathML().length);
					}
                counterMathZone++;
                // append mathzone to paragraph contents
                continue; // move onto next character
                }
            if(insertionPoint.mtIsMath()==2){
                continue; // move onto next character
                // do nothing
                }
            if(insertionPoint.mtIsMath()==3){
                //if character is apart of a mathML segment do nothing as the character was already inserted within the mathML element
                continue; // move onto next character
                }
            if(chr.allGraphics.length>0){
                paragraphContents += addInlineGraphic (chr.allGraphics[0]);
                continue;
                }
			
			if(!chr.hasOwnProperty("appliedCharacterStyle")||!chr.appliedFont.hasOwnProperty("appliedFont")){ // catch arial narrow bold text
				paragraphContents += chrContents;
				continue;
				}
			
            if(chr.appliedFont.name=="Arial\tNarrow Bold"){  // arial narrow bold does not have a fontStyleName property, so just add contents
                    paragraphContents += chrContents;
                    continue;
                    };
				
            // is character bold && statetrue:close false:open bold tag
            if(chr.appliedCharacterStyle.name=="Bold"||chr.appliedFont.fontStyleName=="Bold"){ // character is bold
                if(stateBold==true){
                    paragraphContents += chrContents; // add character to paragraph contents
                    continue;
                    }
                else{
                    paragraphContents += "<bold>"+chrContents; // open bold tag and append character
                    stateBold = true;
                    continue;
                    }
                }
            else //character is not bold
            {
                if(stateBold==true) // if bold tag is active then prepend character with bold tag
                {
                    paragraphContents += "</bold>"; // close bold tag
                    stateBold = false;
                    } 
                }
            
            if(chr.appliedCharacterStyle.name=="Italics"||chr.appliedFont.fontStyleName=="Italic"){ // character is italic
                if(stateItalic==true){
                    paragraphContents += chrContents; // add character to paragraph contents
                    continue;
                    }
                else{
                    paragraphContents += "<italic>"+chrContents; // prepend character with italic tag
                    stateItalic = true;
                    continue;
                    }
                }
            else  // character is not italic
            {
                if(stateItalic==true) // if italic tag is active then close previous italic tag
                {
                    paragraphContents += "</italic>";
                    stateItalic = false;
                    } 
                }
            
            if(chr.position==1936749411){
                paragraphContents += "<sup>"+chrContents+"</sup>";
                continue;
                }
            
            if(chr.position==1935831907){
                paragraphContents += "<sub>"+chrContents+"</sub>";
                continue;
                }
			
			if(chrContents.constructor.name === "Enumerator"){ // page breaks, column breaks
				$.writeln("Enumerator loop reached.");
				if(chrContents==SpecialCharacters.BULLET_CHARACTER){paragraphContents += "&#8226;"}
				else if(chrContents==SpecialCharacters.PAGE_BREAK){}
				else if(chrContents==SpecialCharacters.DOUBLE_LEFT_QUOTE){paragraphContents += "&#8220;"}
				else if(chrContents==SpecialCharacters.DOUBLE_RIGHT_QUOTE){paragraphContents += "&#8221;"}
				else if(chrContents==SpecialCharacters.SINGLE_LEFT_QUOTE){paragraphContents += "&#8216;"}
				else if(chrContents==SpecialCharacters.SINGLE_RIGHT_QUOTE){paragraphContents += "&#8217;"}
				else if(chrContents==SpecialCharacters.COLUMN_BREAK){}
				else if(chrContents==SpecialCharacters.FORCED_LINE_BREAK){}
				}
            paragraphContents += chrContents; // handle regular character
            } //characters
        
        if(stateBold==true){
            paragraphContents += "</bold>"; // if get contents reaches this far then the final character is not bold but the bold is activated
            }
		
        if(stateItalic==true){
            paragraphContents += "</italic>"; // if get contents reaches this far then the final character is not italic but the italic is activated
            }

        var figureText = RegExp("Figure [\\d]+").exec(paragraphContents);
        var tableText = RegExp("Table [\\d]+").exec(paragraphContents);
        if(figureText!=null){
            var reftype = "fig";
            var rid = pad(RegExp("[\\d]+").exec(figureText),3);
            //alert(figureText+", "+rid);
            paragraphContents = paragraphContents.replace(figureText,getRef(reftype,rid,figureText));
            };
        if(tableText!=null){
            var reftype = "table";
            var rid = pad(RegExp("[\\d]+").exec(figureText),3);
            //alert(figureText+", "+rid);
            paragraphContents = paragraphContents.replace(figureText,getRef(reftype,rid,figureText));
            }
        return paragraphContents;
        } //if
    else{
        paragraphContents = p.contents;
        var figureText = RegExp("Figure [\\d]+").exec(paragraphContents);
        var tableText = RegExp("Table [\\d]+").exec(paragraphContents);
        if(figureText!=null){
            var reftype = "fig";
            var rid = pad(RegExp("[\\d]+").exec(figureText),3);
            //alert(figureText+", "+rid);
            paragraphContents = paragraphContents.replace(figureText,getRef(reftype,rid,figureText));
            };
        if(tableText!=null){
            var reftype = "table";
            var rid = pad(RegExp("[\\d]+").exec(figureText),3);
            //alert(figureText+", "+rid);
            paragraphContents = paragraphContents.replace(figureText,getRef(reftype,rid,figureText));
            }
        return paragraphContents;
        }

    var figureText = RegExp("Figure [\\d]+").exec(paragraphContents);
    var tableText = RegExp("Table [\\d]+").exec(paragraphContents);
    if(figureText!=null){
        var reftype = "fig";
        var rid = pad(RegExp("[\\d]+").exec(figureText),3);
        //alert(figureText+", "+rid);
        paragraphContents = paragraphContents.replace(figureText,getRef(reftype,rid,figureText));
        };
    if(tableText!=null){
        var reftype = "table";
        var rid = pad(RegExp("[\\d]+").exec(figureText),3);
        //alert(figureText+", "+rid);
        paragraphContents = paragraphContents.replace(figureText,getRef(reftype,rid,figureText));
        }
    }

function parseParagraph(){
	$.writeln("entering: paragraphParser");
	var paragraphs = sequenceParagraphs (app.activeDocument);
	for(var i = 0; i<paragraphs.length; i++){
		if(paragraphs[i].type == "paragraph"){
			var p = app.activeDocument.pages[paragraphs[i]["pgId"]].textFrames[paragraphs[i]["tfId"]].paragraphs[paragraphs[i]["pId"]];
			if(p.characters.length==0){continue};
			handleParagraph(p);
			}
		else if(paragraphs[i].type == "table"){
			var p = app.activeDocument.pages[paragraphs[i]["pgId"]].textFrames[paragraphs[i]["tfId"]].paragraphs[paragraphs[i]["pId"]];
			handleTable(p);
			}
		else if(paragraphs[i].type == "rectangle"){
			handleRectangle(app.activeDocument.pages[paragraphs[i]["pgId"]].rectangles[paragraphs[i]["rectId"]]);
			}
		}
	}

function classifyParagraph(paragraph){
	$.writeln("entering: classifyParagraph");
	if(paragraph.tables.length>0){
		return "table";
		}
	else
	{
		return paragraph.appliedParagraphStyle.name
		}
	}

function handleOther(){
	$.writeln("entering: handleOther");
	}

function handleParagraph(paragraph){
	$.writeln("entering: handleParagraph");
	$.writeln(paragraph.contents);
	var headingStyles = ["Heading 1","Heading 2","Heading 3","Heading 4","Q & A Header"];
	var bodyStyles = ["Body Text","Body Text - Indented"];
	var listStyles = ['Bullets', 'Number/Letter Indent', 'Number/Letter Indent2', 'Note','Number/Letter Sub', 'Number/Letter Indent_Sub', 'Bullets-sub', 'Q & A Indent'];	
	if(headingStyles.indexOf(paragraph.appliedParagraphStyle.name)!=-1){ // is heading style
		activeList = false;
		iList1 = 0; iList2 = 0; // reset list counters
		handleHeading(paragraph);
		}
	else if(bodyStyles.indexOf(paragraph.appliedParagraphStyle.name)!=-1){ // is body style
		activeList = false;
		iList1 = 0; iList2 = 0; // reset list counters
		handleBodyParagraph(paragraph);
		}
	else if(listStyles.indexOf(paragraph.appliedParagraphStyle.name)!=-1){ // is list style
		handleList(paragraph);
		}
	else{
		missingStyles[missingStyles.length] = paragraph.appliedParagraphStyle.name;
		}
	}

function handleTable(p){
	$.writeln("entering: handleTable");
	var tblClass = classifyTable(p);
	if(tblClass == "tables in table"){
		for(var i = 0; i<p.tables[0].cells.length; i++){
			handleTable(p.tables[0].cells[i].paragraphs[0]);
			}
		}
	if(tblClass == "TFigure Header"){
		if(typeof iFigure === 'undefined'){iFigure = 1}
		else{iFigure++}
		
		p.justification = Justification.CENTER_ALIGN; // ensure justified graphic to center
		tblToJPG (p.tables[0],p,"figure",iFigure); // create image
		getActiveHeadingNode ().appendChild(new XML("<fig>"));
		getActiveHeadingNode ().fig[getActiveHeadingNode ().fig.length()-1].@['fig-type'] = "figure";
		getActiveHeadingNode ().fig[getActiveHeadingNode ().fig.length()-1].@['id'] = "CH"+pad(iChapter,3)+".FIG"+pad(iFigure,3);
		getActiveHeadingNode ().fig[getActiveHeadingNode ().fig.length()-1].appendChild(new XML("<p>"));
		var activeP = getActiveHeadingNode ().fig[getActiveHeadingNode ().fig.length()-1].p[getActiveHeadingNode ().fig[getActiveHeadingNode ().fig.length()-1].p.length()-1];
		activeP['inline-graphic'][activeP['inline-graphic'].length()-1].@['mime-subtype'] = "jpg";
		activeP['inline-graphic'][activeP['inline-graphic'].length()-1].@['xlink::href'] = "CH"+pad(iChapter,3)+".FIG"+pad(iFigure,3);
		activeP['inline-graphic'][activeP['inline-graphic'].length()-1].@['xlink::title'] = "Figure "+iFigure;
		activeP['inline-graphic'][activeP['inline-graphic'].length()-1].@['xhtml::xlink'] = "http://www.w3.org/1999/xlink";
		manifest[manifest.length] = {href:"CH"+pad(iChapter,3)+"."+"FIG"+pad(iFigure+1,3)};
		}
	else if(tblClass == "Q & A Header"){
		handleHeaderTable(tblClass,p);
		}
	else if(tblClass == "Objective Header"){
		var table = p.tables[0];
		handleObjectiveTable (table);
		}
	else if(tblClass == "Chapter LO - Header"){
		var table = p.tables[0];
		handleOutcomeTable (table);
		}
	else if(tblClass == "Table Header"){
		handleTableHeader(p);
		}
	else if(tblClass == "other"){
		handleMiscTable(p);
		$.writeln("other table class");
		}
	else if(typeof tblClass === 'undefined'){
		$.writeln("undefined table class");
		}
		
	}

function handleHeading(heading){
	$.writeln("entering: handleHeader");
	var headingLevel = setHeadingLevel(heading);
	}


function handleBodyParagraph(bodyParagraph){
	$.writeln("entering: handleBodyParagraph");
	if(tmpHeadingLevel == 1){
		chapter.sec[iH1-1].appendChild(new XML("<p>"));
		chapter.sec[iH1-1].p[chapter.sec[iH1-1].p.length()-1] = getContents (bodyParagraph);
		};
	else if(tmpHeadingLevel == 2){
		chapter.sec[iH1-1].sec[iH2-1].appendChild(new XML("<p>"));
		chapter.sec[iH1-1].sec[iH2-1].p[chapter.sec[iH1-1].sec[iH2-1].p.length()-1] = getContents (bodyParagraph);
		};
	else if(tmpHeadingLevel == 3){
		chapter.sec[iH1-1].sec[iH2-1].sec[iH3-1].appendChild(new XML("<p>"));
		chapter.sec[iH1-1].sec[iH2-1].sec[iH3-1].p[chapter.sec[iH1-1].sec[iH2-1].sec[iH3-1].p.length()-1] = getContents (bodyParagraph);
		};
	else if(tmpHeadingLevel == 4){
		chapter.sec[iH1-1].sec[iH2-1].sec[iH3-1].sec[iH4-1].appendChild(new XML("<p>"));
		chapter.sec[iH1-1].sec[iH2-1].sec[iH3-1].sec[iH4-1].p[chapter.sec[iH1-1].sec[iH2-1].sec[iH3-1].sec[iH4-1].p.length()-1] = getContents (bodyParagraph);
		};
	else{}
	}

function setHeadingLevel(heading){
	iList1 = 0; iList2 = 0;
	$.writeln("entering: setHeadingLevel");
	var style = heading.appliedParagraphStyle.name;
	if(typeof tmpHeadingLevel === 'undefined'){tmpHeadingLevel = 0};
	var headingLevel = style.substring(0,9).replace("Heading ","");
	if(style == "Q & A Header"){headingLevel = 1};
	if(headingLevel == 1){
		iH2 = 0;
		iH3 = 0;
		iH4 = 0;
		if(typeof iH1 === 'undefined'){iH1 = 0}
		if(heading.contents == "Glossary\r"){
			}
		tmpHeadingLevel = 1;
		chapter.appendChild(new XML("<sec>"));
		iH1++;
		chapter.sec[iH1-1].@['disp-level'] = headingLevel;
		chapter.sec[iH1-1].@id = "CH"+pad(iChapter,3)+"."+"SEC"+pad(chapter.sec.length(),3);
		chapter.sec[iH1-1].@['sec-type'] = "section";
		chapter.sec[iH1-1].title = heading.contents;
		}
	else if(headingLevel == 2){
		if(iH1 == 0){chapter.appendChild(new XML("<sec>")); iH1++};
		iH3 = 0;
		iH4 = 0;
		if(typeof iH2 === 'undefined'){iH2 = 0}
		tmpHeadingLevel = 2;
		chapter.sec[iH1-1].appendChild(new XML("<sec>"));
		iH2++;
		chapter.sec[iH1-1].sec[iH2-1].@['disp-level'] = headingLevel;
		chapter.sec[iH1-1].sec[iH2-1].@id = "CH"+pad(iChapter,3)+"."+"SEC"+pad(chapter.sec.length(),3)+"."+iH2;
		chapter.sec[iH1-1].sec[iH2-1].@['sec-type'] = "section";
		chapter.sec[iH1-1].sec[iH2-1].title = heading.contents;
		}
	else if(headingLevel == 3){
		if(iH1 == 0){chapter.appendChild(new XML("<sec>")); iH1++};
		if(iH2 == 0){chapter.sec[iH1-1].appendChild(new XML("<sec>")); iH2++};
		iH4 = 0;
		if(typeof iH3 === 'undefined'){iH3 = 0}
		tmpHeadingLevel = 3;
		chapter.sec[iH1-1].sec[iH2-1].appendChild(new XML("<sec>"));
		iH3++;
		chapter.sec[iH1-1].sec[iH2-1].sec[iH3-1].@['disp-level'] = headingLevel;
		chapter.sec[iH1-1].sec[iH2-1].sec[iH3-1].@id = "CH"+pad(iChapter,3)+"."+"SEC"+pad(chapter.sec.length(),3)+"."+iH2+"."+iH3;
		chapter.sec[iH1-1].sec[iH2-1].sec[iH3-1].@['sec-type'] = "section";
		chapter.sec[iH1-1].sec[iH2-1].sec[iH3-1].title = heading.contents;
		}
	else if(headingLevel == 4){
		if(iH1 == 0){chapter.appendChild(new XML("<sec>")); iH1++};
		if(iH2 == 0){chapter.sec[iH1-1].appendChild(new XML("<sec>")); iH2++};
		if(iH3 == 0){chapter.sec[iH1-1].sec[iH2-1].appendChild(new XML("<sec>")); iH3++};
		if(typeof iH4 === 'undefined'){iH4 = 0}	
		tmpHeadingLevel = 4;
		chapter.sec[iH1-1].sec[iH2-1].sec[iH3-1].appendChild(new XML("<sec>"));
		iH4++;
		chapter.sec[iH1-1].sec[iH2-1].sec[iH3-1].sec[iH4-1].@['disp-level'] = headingLevel;
		chapter.sec[iH1-1].sec[iH2-1].sec[iH3-1].sec[iH4-1].@id = "CH"+pad(iChapter,3)+"."+"SEC"+pad(chapter.sec.length(),3)+"."+iH2+"."+iH3+"."+iH4;
		chapter.sec[iH1-1].sec[iH2-1].sec[iH3-1].sec[iH4-1].@['sec-type'] = "section";
		chapter.sec[iH1-1].sec[iH2-1].sec[iH3-1].sec[iH4-1].title = heading.contents;
		}
	else{
		alert("Unmapped heading.");
		}
	//alert(String(heading.contents+", "+iH1+", "+iH2+", "+iH3+", "+iH4))
	}

function getRef(reftype, rid, str){
    return "<xref ref-type=\""+reftype+"\" rid=\""+"CH"+pad(iChapter,3)+"."+String(reftype.substring(0,3).toUpperCase())+rid+"\">"+str+"</xref>";
    }

function handleList(listParagraph){
	var arrListStyles = ['Bullets', 'Number/Letter Indent', 'Number/Letter Indent2', 'Note'];
	var arrSublistStyles = ['Number/Letter Sub', 'Number/Letter Indent_Sub', 'Bullets-sub', 'Q & A Indent'];
	if(typeof iList1 === 'undefined'){iList1 = 0};
	if(typeof iList2 === 'undefined'){iList2 = 0};
	if(arrListStyles.indexOf(listParagraph.appliedParagraphStyle.name) != -1){ // is list
		iList2 = 0;
		if(activeList == false){
			getActiveHeadingNode ().appendChild(new XML("<list>"));
			if(listParagraph.appliedParagraphStyle.name=="Bullets"){
				getActiveListNode ().@['list-type']="bullet";
				}
			else{
				getActiveListNode ().@['list-type']="simple";
				}
			activeList = true;
			}
		getActiveListNode ().appendChild(new XML("<listitem>"));
		getActiveListNode ().listitem[getActiveListNode ().listitem.length()-1].appendChild(new XML("<p>"));
		getActiveListNode ().listitem[getActiveListNode ().listitem.length()-1].p[0] = listParagraph.contents;
		}
	else if(arrSublistStyles.indexOf(listParagraph.appliedParagraphStyle.name) != -1){ // is sublist
		if(iList2 == 0){
			getActiveListNode ().listitem[getActiveListNode ().listitem.length()-1].appendChild(new XML("<list>"));
			if(listParagraph.appliedParagraphStyle.name=="Bullets"){
				getActiveListNode ().@['list-type']="bullet";
				}
			else{
				$.writeln(iH1+", "+iH2+", "+iH3+", "+iH4);
				$.writeln(iList1+", "+iList2);
				getActiveListNode ().@['list-type']="simple";
				}
			iList2++;
			}
		getActiveListNode ().appendChild(new XML("<listitem>"));
		getActiveListNode ().listitem[getActiveListNode ().listitem.length()-1].appendChild(new XML("<p>"));
		getActiveListNode ().listitem[getActiveListNode ().listitem.length()-1].p[0] = listParagraph.contents;
		}
	else
	{
		alert("Error: Not a listed list style. "+listParagraph.appliedParagraphStyle.name);
		}
	}

function getActiveHeadingNode(){
	if(tmpHeadingLevel == 1){
		return chapter.sec[iH1-1];
		}
	else if(tmpHeadingLevel == 2){
		return chapter.sec[iH1-1].sec[iH2-1];
		}
	else if(tmpHeadingLevel == 3){
		return chapter.sec[iH1-1].sec[iH2-1].sec[iH3-1];
		}
	else if(tmpHeadingLevel == 4){
		return chapter.sec[iH1-1].sec[iH2-1].sec[iH3-1].sec[iH4-1];
		}
	}

function getActiveListNode(){
	// if iList2 not active
	if(iList2 == 0){
		if(iList1 == 0){getActiveHeadingNode().appendChild(new XML("<list>")); iList1++};
		return getActiveHeadingNode().list[getActiveHeadingNode().list.length()-1];
		}
	else if(iList2 > 0){
		return getActiveHeadingNode().list[getActiveHeadingNode().list.length()-1].listitem[getActiveHeadingNode().list[getActiveHeadingNode().list.length()-1].listitem.length()-1].list[getActiveHeadingNode().list[getActiveHeadingNode().list.length()-1].listitem[getActiveHeadingNode().list[getActiveHeadingNode().list.length()-1].listitem.length()-1].list.length()-1];
		}
	}

function classifyTable(p){
	$.writeln("entering: classifyTable");
	if(p.tables.length > 1){
		return;
		}
	else if(p.tables.length == 1){
		$.writeln("table length is 1");
		var table = p.tables[0];
		for(var i = 0; i<table.cells.length; i++){
			var cell = table.cells[i];
			if(cell.tables.length > 0){
				//alert("table contains tables")
				return "tables in table";
				}
			else{
				$.writeln("entering cell loop");
				for(var j = 0; j<cell.paragraphs.length; j++){
					var p = cell.paragraphs[j];
					if(p.appliedParagraphStyle.name == "TFigure Header"||p.appliedParagraphStyle.name == "Q & A Header"||p.appliedParagraphStyle.name == "Objective Header"||p.appliedParagraphStyle.name == "Chapter LO - Header"||p.appliedParagraphStyle.name == "Table Header"){
						$.writeln(p.appliedParagraphStyle.name);
						return p.appliedParagraphStyle.name;
						}
					$.writeln("returning null");
					return;
					}
				}
			}
		return "other";
		}
	else{
		//alert("paragraph does not contain tables")
		return;
		}
	
	}

function handleHeaderTable(type,p){
	if(p.tables.length == 1){
		var table = p.tables[0];
		}
	for(var i = 0; i<table.cells.length; i++){
		var cell = table.cells[i];
		for(var j = 0; j<table.cells[i].paragraphs.length; j++){
			var p = table.cells[i].paragraphs[j];
			if(p.appliedParagraphStyle.name == "Q & A Header"){
				handleParagraph(p);
				return;
				}
			}
		}
	}

function tblToJPG(table,parentParagraph,tblClass,varCount){
	$.writeln("entering: tblToJPG");
	switch(tblClass){
		case "table":
		tblClassCode = "TAB";
		break;
		case "figure":
		tblClassCode = "FIG";
		break;
		case "other":
		tblClassCode = "MTAB";
		break;
		default:
		alert("Problem encountered in function tblToJPG. The table class "+tblClass+" is unknown.");
		}
	app.select(table);
	$.writeln("table selected");
	app.copy();
	$.writeln("entering: tblToJPG:copy");
	var docName = app.activeDocument.name;
	app.activeDocument.zeroPoint=[0,0];
	app.activeDocument.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.POINTS;
	app.activeDocument.viewPreferences.verticalMeasurementUnits = MeasurementUnits.POINTS;
	app.viewPreferences.verticalMeasurementUnits = MeasurementUnits.POINTS;
	app.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.POINTS;
	var newDoc = app.documents.add(true,{
		documentPreferences:{
			pageHeight:table.height*1.05+"pt",
			pageWidth:table.width*1.05+"pt",
			createPrimaryTextFrame:true,
			viewPreferences:{
				horizontalMeasurementUnits:MeasurementUnits.POINTS,
				verticalMeasurementUnits:MeasurementUnits.POINTS
				}
			}
		});
	$.writeln("table added");
	newDoc.pages[0].textFrames[0].geometricBounds = [0,0,newDoc.documentPreferences.pageHeight,newDoc.documentPreferences.pageWidth];
	app.select(newDoc.pages[0].textFrames[0].insertionPoints[0])
	app.paste();
	$.writeln("table pasted");
	app.activeDocument.pages[0].textFrames[0].textFramePreferences.verticalJustification = VerticalJustification.CENTER_ALIGN;
	app.activeDocument.pages[0].textFrames[0].paragraphs[0].justification = Justification.CENTER_ALIGN;
	var chapterFolder = "C:\\Users\\kstaples\\Desktop\\0001 - Projects\\47 - New NLM\\NLM Out\\"+docName.replace(".indd","")+"\\";
	if(!Folder(chapterFolder).exists){
		Folder(chapterFolder).create();
		}
	app.jpegExportPreferences.exportResolution = 96;
	//newDoc.save("c:\\users\\kstaples\\desktop\\figuredocuments\\"+docName.replace(".indd","")+"."+tblClassCode+pad(varCount,3)+".indd");
	newDoc.exportFile("JPEG",chapterFolder+docName.replace(".indd","")+"."+tblClassCode+pad(varCount,3)+".jpg",false);
	newDoc.close(SaveOptions.NO);
	$.writeln(app.activeDocument.name);
	//app.activeDocument = app.documents.itemByName(docName);
	}// function

function handleOutcomeTable(table){
    // add header outcome to document; chapter/paragraph/boxedtext/
    var headerOutcome = table.cells[0];
    chapter.appendChild(new XML("<p>")) // add paragraph to root
    //adding learning outcomes to document
	var activeP = chapter.p[chapter.p.length()-1];
	activeP.appendChild(new XML("<boxed-text>"));
	if(typeof iBoxedText === 'undefined'){iBoxedText = 1}else{iBoxedText++}
    activeP['boxed-text'][activeP['boxed-text'].length()-1].@id = "CH"+pad(iChapter,3)+"."+"BOX"+pad(iBoxedText,3); //set boxed text properties
	activeP['boxed-text'][activeP['boxed-text'].length()-1].@['content-type'] = "gray";
    var title = table.cells[0].paragraphs[0].contents; //learning objectives title
    activeP['boxed-text'][activeP['boxed-text'].length()-1].title.italic = title;
	var objectiveBox = activeP['boxed-text'][activeP['boxed-text'].length()-1];
    var listOutcomes = table.cells[1];
    for(var i = 0; i < listOutcomes.paragraphs.length; i++){
        objectiveBox.appendChild(new XML("<p>"));
        var activeP = objectiveBox.p[objectiveBox.p.length()-1]; // get sub paragraph index
        activeP.bold.italic = listOutcomes.paragraphs[i].contents;
        }
	
    var headerObjectives = table.cells[2];
	chapter.appendChild(new XML("<p>")) // add paragraph to root
	var activeP = chapter.p[chapter.p.length()-1];
    activeP.appendChild(new XML("<boxed-text>")); // add boxed text
    if(typeof iBoxedText === 'undefined'){iBoxedText = 1}else{iBoxedText++}
	activeP['boxed-text'][activeP['boxed-text'].length()-1].@id = "CH"+pad(iChapter,3)+"."+"BOX"+pad(iBoxedText,3); //set boxed text properties
	activeP['boxed-text'][activeP['boxed-text'].length()-1].@['content-type'] = "gray";
    var title = headerObjectives.paragraphs[0].contents; //learning objectives title
    activeP['boxed-text'][activeP['boxed-text'].length()-1].title.italic = title;
    var listObjectives = table.cells[3];
    for(var i = 0; i<listObjectives.paragraphs.length; i++){ // for each paragraph in objective list
        var paragraph = listObjectives.paragraphs[i];
        if(paragraph.appliedParagraphStyle.name == "Chapter LO - First Line" && paragraph.contents.length > 1){
           activeP['boxed-text'][activeP['boxed-text'].length()-1].appendChild(new XML("<p>"));
		   var activeP = activeP['boxed-text'][activeP['boxed-text'].p.length()-1];
		   activeP.bold.italic = paragraph.contents;
            };
        if(paragraph.appliedParagraphStyle.name == "Chapter LO - Numbered" || paragraph.appliedParagraphStyle.name == "Chapter Text - Numbered" || paragraph.appliedParagraphStyle.name == "Chapter LO - TOC list"){
			if(activeP.list.length()==0){
				activeP.appendChild(new XML("<list>"));
				var activeList = activeP.list[activeP.list.length()-1];
				activeList.@['list-type'] = "simple";
				}
			activeList.appendChild(new XML("<list-item>"));
			activeListItem = activeList['list-item'][activeList['list-item'].length()-1];
			activeListItem.appendChild(new XML("<p>"));
			var activeParagraph = activeListItem.p[activeListItem.p.length()-1];
			activeParagraph.bold.italic = paragraph.bulletsAndNumberingResultText+getContents(paragraph);
            }
        } 
    }

function handleObjectiveTable(table){
    // add header outcome to document; chapter/paragraph/boxedtext/
    var headerObjective = table.cells[0];
	getActiveHeadingNode ().appendChild(new XML("<p>"));
	var activeP = getActiveHeadingNode ().p[getActiveHeadingNode ().p.length()-1];
	activeP.appendChild(new XML("<boxed-text>")); // add boxed text
	if(typeof iBoxedText === 'undefined'){iBoxedText = 1}else{iBoxedText++};
    activeP['boxed-text'][activeP['boxed-text'].length()-1].@id = "CH"+pad(iChapter,3)+"."+"BOX"+pad(iBoxedText,3); //set boxed text properties
    activeP['boxed-text'][activeP['boxed-text'].length()-1].@['content-type'] ="gray"; // set properties
    var title = headerObjective.paragraphs[0].contents; //learning objective title
	activeP['boxed-text'][activeP['boxed-text'].length()-1].title.italic = title;
    var txtObjective = table.cells[1];
    for(var i = 0; i<txtObjective.paragraphs.length; i++){
        activeP['boxed-text'][activeP['boxed-text'].length()-1].appendChild(new XML("<p>"));
        activeP['boxed-text'][activeP['boxed-text'].length()-1].p[activeP['boxed-text'][activeP['boxed-text'].length()-1].p.length()-1] = txtObjective.paragraphs[i].contents;
        }
    }

function createManifest(fManifest){
	//create chapter folder
    var chapterFolder = "C:\\Users\\kstaples\\Desktop\\0001 - Projects\\47 - New NLM\\NLM OUT\\"+app.activeDocument.name.replace(".indd","")+"\\";
    if(Folder(chapterFolder).exists==false){
        try{Folder(chapterFolder).create()}catch(e){alert(e)};
        }
        var xmlManifest = new XML("<files>")
        //add chapter file to manifest
        xmlManifest.appendChild(new XML("<file>"));
        xmlManifest.file[xmlManifest.file.length()-1] = app.activeDocument.name.replace("indd","xml");
        fManifest.sort(function (a,b){return a.href>b.href})
    for(var i = 0; i<fManifest.length; i++){
        xmlManifest.appendChild(new XML("<file>"));
        xmlManifest.file[xmlManifest.file.length()-1] = fManifest[i].href+".jpg";
        }//i
	fleManifest = File(chapterFolder+"manifest.xml");
    fleManifest.open('w');
    fleManifest.encoding = 'UTF-8';
    fleManifest.write(xmlManifest);
    fleManifest.close();
    }

function sequenceParagraphs(doc) {
	var paragraphs = [];
	//process document by spreads, pages, textFrames, paragraphs, tables
	for (var i = 0; i < doc.pages.length; i++) {
		var arrListStyles = ['Bullets', 'Number/Letter Indent', 'Number/Letter Indent2', 'Note'];
		var arrSublistStyles = ['Number/Letter Sub', 'Number/Letter Indent_Sub', 'Bullets-sub', 'Q & A Indent'];
		var page = app.activeDocument.pages[i];
		for(var k = 0; k < page.rectangles.length; k++){
			paragraphs[paragraphs.length] = {
				type: "rectangle",
				rectId: k,
				pgId: i,
				gb: getRectangleBounds (page.rectangles[k])
				}
			}
		for (var k = 0; k < page.textFrames.length; k++) {
			for (var l = 0; l < page.textFrames[k].paragraphs.length; l++) {
				var p = page.textFrames[k].paragraphs[l];
				if(p.contents.replace(/\s/gi,"").length==0){ // skip blank paragraphs
					continue;
					}
				else if(p.contents.indexOf('\uFFFC') > -1){ // skip anchored objects
					continue;
					}
				if(p.contents.indexOf('\u0016') != -1 && p.tables.length == 1) { // HOW TO HANDLE PARAGRAPHS WHICH ARE STRITCLY TABLES
					if(typeof tblID === "undefined"){
						tblID = p.tables[0].id;
						}
					else if(tblID == p.tables[0].id){
						// do nothing - duplicate table
						}
					else{
						tblID = p.tables[0].id;
						paragraphs[paragraphs.length] = {
							type: "table",
							pgId: i,
							tfId: k,
							pId: l,
							gb: getBounds(p)
							};
						//$.writeln("RECTANGLE:") //+p.contents+", "+pgParagraphs[pgParagraphs.length-1].gb.toString());
						}
					}
				else {
					paragraphs[paragraphs.length] = {
						type: "paragraph",
						pgId: i,
						tfId: k,
						pId: l,
						gb: getBounds(p)
						};
					//$.writeln("PARAGRAPH:") //+p.contents+", "+pgParagraphs[pgParagraphs.length-1].gb.toString());
					}
				}
			}
		}
	paragraphs.sort(function(a,b){return a["pgId"] - b["pgId"] || a["gb"][2] - b["gb"][2];});
	return paragraphs;
	}

function getBounds(p){
	if(p.characters[0].tables.length>0){
		return getTableBounds(p);
		}
	else
	{
		return getParagraphBounds(p);
		}
	}

function getTableBounds(tbl){
	app.select(tbl);
	app.activeDocument.zeroPoint=[0,0];
	app.activeDocument.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.POINTS;
	app.activeDocument.viewPreferences.verticalMeasurementUnits = MeasurementUnits.POINTS;
	var y1 = app.selection[0].insertionPoints[0].baseline-app.selection[0].tables[0].height;
	var x1 = app.selection[0].insertionPoints[0].horizontalOffset;
	//if(x1>613){x1 -= 613};
	var y2 = app.selection[0].insertionPoints[0].baseline;
	var x2 = app.selection[0].insertionPoints[0].horizontalOffset+app.selection[0].tables[0].width;
	//if(x2>613){x2 -= 613};
	var GB = [y1,x1,y2,x2];
	return GB;
	}

function getParagraphBounds(p){
	app.activeDocument.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.POINTS;
	app.activeDocument.viewPreferences.verticalMeasurementUnits = MeasurementUnits.POINTS;
	app.select(p);
	//$.writeln(p.constructor.name+": "+p.contents+"\rlength:"+p.contents.length);
	if(p.contents=="\r"){return false};
	if(p.contents.indexOf('\FFFC') > -1){return false}
	var o = p.createOutlines(false);
	var GB = (o[0].geometricBounds);
	//app.activeDocument.undo();
	return GB;
	}

function getRectangleBounds(r){
	app.activeDocument.zeroPoint=[0,0];
	app.activeDocument.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.POINTS;
	app.activeDocument.viewPreferences.verticalMeasurementUnits = MeasurementUnits.POINTS;
	var GB = r.geometricBounds;
	return GB;
	}

function handleRectangle(r){
	$.writeln("entering: handleRectangle");
	if(r.allGraphics.length > 0){ // contains graphics
		//contains media elements
		if(r.allGraphics.length==1){
			var g = r.allGraphics[0]; // graphic
			if(g.itemLink.filePath.indexOf("Media Element") != -1); // is media element
			{
				var URL = getLinkURL (r);
				getActiveHeadingNode ().appendChild(new XML("<p>")); // create paragraph for media element
				var activeParagraph = getActiveParagraph ();
				addMediaToObject(activeParagraph,getLinkURL (r));
				addGraphicToObject (activeParagraph.media[activeParagraph.media.length()-1], r.allGraphics[0]);
				/*
				getActiveHeadingNode ().appendChild(new XML("<p>"));
				getActiveParagraph ().appendChild(new XML("<media>"));
				getActiveParagraph ().media[getActiveParagraph ().media.length()-1].@['xlink::href'] = URL;
				getActiveParagraph ().media[getActiveParagraph ().media.length()-1].appendChild(new XML("<inline-graphic>"));
				getActiveParagraph ().media[getActiveParagraph ().media.length()-1]["inline-graphic"].@['xlink::href'] = "";
				getActiveParagraph ().media[getActiveParagraph ().media.length()-1]["inline-graphic"].@['xlink::mime-subtype'] = "";
				//alert(g.itemLink.hyperlinks.length);
				//handleMediaElement(g.itemLink);
				*/
				}
			}
		else{
			alert("error thrown in function handleRectangle");
			}
		}
	}

function getLinkURL(objSource){
	var hyperlinks = app.activeDocument.hyperlinks;
	for(var i = 0; i<hyperlinks.length; i++){
		var hyperlink = hyperlinks[i];
		var source = hyperlink.source.sourcePageItem;
		if(source == objSource){
			return hyperlink.destination.destinationURL;
			}
		}
	}

function getActiveParagraph(){
	$.writeln("entering: getActiveParagraph");
	$.writeln(iH1+", "+iH2+", "+iH3+", "+iH4);
	if(tmpHeadingLevel == 1){
		var activeP = chapter.sec[iH1-1].p[chapter.sec[iH1-1].p.length()-1];
		$.writeln(chapter.sec[iH1-1].p.length()-1);
		return activeP;
		};
	else if(tmpHeadingLevel == 2){
		var activeP = chapter.sec[iH1-1].sec[iH2-1].p[chapter.sec[iH1-1].sec[iH2-1].p.length()-1];
		$.writeln(chapter.sec[iH1-1].sec[iH2-1].p.length()-1);
		return activeP;
		};
	else if(tmpHeadingLevel == 3){
		var activeP = chapter.sec[iH1-1].sec[iH2-1].sec[iH3-1].p[chapter.sec[iH1-1].sec[iH2-1].sec[iH3-1].p.length()-1];
		$.writeln(chapter.sec[iH1-1].sec[iH2-1].sec[iH3-1].p.length()-1);
		return activeP;
		};
	else if(tmpHeadingLevel == 4){
		var activeP = chapter.sec[iH1-1].sec[iH2-1].sec[iH3-1].sec[iH4-1].p[chapter.sec[iH1-1].sec[iH2-1].sec[iH3-1].sec[iH4-1].p.length()-1];
		$.writeln(chapter.sec[iH1-1].sec[iH2-1].sec[iH3-1].sec[iH4-1].p.length()-1);
		return activeP;
		};
	else{alert("error heading level not selected");}
	}

function addMediaToObject(parentObject,URL){
	parentObject.appendChild(new XML("<media>"));
	parentObject.media[parentObject.media.length()-1].@['xlinkhref'] = URL;
	}

function addGraphicToObject(parentObject,graphic){
	parentObject.appendChild(new XML("<inline-graphic>"));
	graphicToJPG(graphic);
	var extension = decodeURI(File(graphic.itemLink.filePath).name.split(".")[File(graphic.itemLink.filePath).name.split(".").length-1]);
	var jpgFilename = decodeURI(File(graphic.itemLink.filePath).name.replace(extension,"jpg"));
	parentObject["inline-graphic"].@['xlinkhref'] = jpgFilename;
	parentObject["inline-graphic"].@['xlinkmime-subtype'] = "jpg";
	}

function getLinkURL(objSource){
	var hyperlinks = app.activeDocument.hyperlinks;
	for(var i = 0; i<hyperlinks.length; i++){
		var hyperlink = hyperlinks[i];
		var source = hyperlink.source.sourcePageItem;
		if(source == objSource){
			return hyperlink.destination.destinationURL;
			}
		}
	}

function graphicToJPG(graphic){
	if(typeof graphic === 'undefined'){alert("Graphic undefined.")};
	var extension = decodeURI(File(graphic.itemLink.filePath).name.split(".")[File(graphic.itemLink.filePath).name.split(".").length-1]);
	var jpgFilename = File(graphic.itemLink.filePath).name.replace(extension,"jpg");
	if(File(chapterFolder+jpgFilename).exists){
		// do nothing
		}
	else{
		// create file
		graphic.exportFile("JPEG",decodeURI (chapterFolder+jpgFilename).replace("\\","\\\\")+"\\",false);
		// check if exists
		if(!File(chapterFolder+jpgFilename).exists){
			alert("File not created successfully: "+jpgFilename);
			}
		manifest[manifest.length] = {href:jpgFilename};
		}
	}

function GetSubFolders(theFolder) {  
     var myFileList = theFolder.getFiles();  
     for (var i = 0; i < myFileList.length; i++) {  
          var myFile = myFileList[i];  
          if (myFile instanceof Folder){  
               GetSubFolders(myFile);  
          }  
          else if (myFile instanceof File && myFile.name.match(/\.indd$/i)) {  
               myFiles.push(myFile);  
          }  
     }  
}

function resetGlobals(){
	iH1 = undefined ();
	iH2 = undefined ();
	iH3 = undefined ();
	iH4 = undefined ();
	iList1 = undefined ();
	iList2 = undefined ();
	iFigure = undefined();
	iTable = undefined();
	iInlineGraphic = undefined();
	itmpHeadingLevel = undefined();
	tblID = undefined();
	manifest = [];
	}

function undefined(){
	return;
	}

function addInlineGraphic(image)
{
    var chapterNumber = pad(iChapter,3);
    var chapterFolder = "C:\\Users\\kstaples\\Desktop\\0001 - Projects\\23 - NLM Conversion\\NLM Out\\"+app.activeDocument.name.replace(".indd","")+"\\";
    if(Folder(chapterFolder).exists){}else{Folder(chapterFolder).create()}
    var chapterNumber=pad(iChapter,3);
    iInlineGraphic++;
    var chapterFolder = "C:\\Users\\kstaples\\Desktop\\0001 - Projects\\23 - NLM Conversion\\NLM Out\\"+app.activeDocument.name.replace(".indd","")+"\\";
    if(Folder(chapterFolder).exists){}else{
        Folder(chapterFolder).create()
        }
    image.exportFile("JPEG",chapterFolder+"CH"+chapterNumber+"."+"INLINE"+pad(iInlineGraphic,3)+".jpg",false);
	if(!File(chapterFolder+"CH"+chapterNumber+"."+"INLINE"+pad(iInlineGraphic,3)+".jpg").exists){
	//$.writeln(chapterFolder+"CH"+chapterNumber+"."+"INLINE"+pad(g.iInlineGraphic,3)+".jpg"+" does not exist.");
	}
    inlineGraphic = new XML("<inlinegraphic>");
    inlineGraphic.@mimesubtype = "jpg";
    inlineGraphic.@xlinkhref = "CH"+chapterNumber+"."+"INLINE"+pad(iInlineGraphic,3);
    inlineGraphic.@xmlns::xlink = """http://www.w3.org/1999/xlink""";
    return inlineGraphic.toXMLString();
    }

function convertBulletsAndNumberingToText(){
	for(var i = app.activeDocument.pages.length-1; i >= 0; i--){
		var pg = app.activeDocument.pages[i];
		for(var j = pg.textFrames.length-1; j >= 0; j--){
			var tf = pg.textFrames[j];
			for(var k = tf.paragraphs.length-1; k >= 0; k--){
				if(tf.paragraphs[k].tables.length > 0){
					for(var l = tf.paragraphs[k].tables.length-1; l >= 0; l--){
						tf.paragraphs[k].tables[l].convertBulletsAndNumberingToText();
						app.select(tf.paragraphs[k].tables[l]);
						for(var m = tf.paragraphs[k].tables[l].cells.length-1; m >= 0; m--){
							if(tf.paragraphs[k].tables[l].cells[m].paragraphs.length>0){
								if(tf.paragraphs[k].tables[l].cells[m].paragraphs[0].tables.length>0){
									for(var o = tf.paragraphs[k].tables[l].cells[m].paragraphs[0].tables.length-1; o >= 0; o--){
										tf.paragraphs[k].tables[l].cells[m].paragraphs[0].tables[o].convertBulletsAndNumberingToText();
										app.select(tf.paragraphs[k].tables[l].cells[m].paragraphs[0].tables[o]);
										//$.sleep(125);
										}
									}
								}
							}
						}
					}
				}
			}
		}
	}

function handleTableHeader(p){
	var table = p.tables[0];
	if(typeof iTable === 'undefined'){iTable = 1}
	else{iTable++}
	p.justification = Justification.CENTER_ALIGN; // ensure justified graphic to center
	tblToJPG (p.tables[0],p,"table",iTable); // create image
	getActiveHeadingNode ().appendChild(new XML("<fig>"));
	getActiveHeadingNode ().fig[getActiveHeadingNode ().fig.length()-1].@['fig-type'] = "figure";
	getActiveHeadingNode ().fig[getActiveHeadingNode ().fig.length()-1].@['id'] = "CH"+pad(iChapter,3)+".TAB"+pad(iTable,3);
	getActiveHeadingNode ().fig[getActiveHeadingNode ().fig.length()-1].appendChild(new XML("<p>"));
	var activeP = getActiveHeadingNode ().fig[getActiveHeadingNode ().fig.length()-1].p[getActiveHeadingNode ().fig[getActiveHeadingNode ().fig.length()-1].p.length()-1];
	activeP['inline-graphic'][activeP['inline-graphic'].length()-1].@['mime-subtype'] = "jpg";
	activeP['inline-graphic'][activeP['inline-graphic'].length()-1].@['xlink::href'] = "CH"+pad(iChapter,3)+".TAB"+pad(iTable,3);
	activeP['inline-graphic'][activeP['inline-graphic'].length()-1].@['xlink::title'] = "Table "+iTable;
	activeP['inline-graphic'][activeP['inline-graphic'].length()-1].@['xhtml::xlink'] = "http://www.w3.org/1999/xlink";
	manifest[manifest.length] = {href:"CH"+pad(iChapter,3)+"."+"TAB"+pad(iTable+1,3)};
	}

function handleMiscTable(p){
	var table = p.tables[0];
	//handle table header
	if(typeof iMiscTable === 'undefined'){iMiscTable = 1}
	else{iMiscTable++}
	p.justification = Justification.CENTER_ALIGN; // ensure justified graphic to center
	tblToJPG (p.tables[0],p,"table",iTable); // create image
	getActiveHeadingNode ().appendChild(new XML("<fig>"));
	getActiveHeadingNode ().fig[getActiveHeadingNode ().fig.length()-1].@['fig-type'] = "figure";
	getActiveHeadingNode ().fig[getActiveHeadingNode ().fig.length()-1].@['id'] = "CH"+pad(iChapter,3)+".MTAB"+pad(iMiscTable,3);
	getActiveHeadingNode ().fig[getActiveHeadingNode ().fig.length()-1].appendChild(new XML("<p>"));
	var activeP = getActiveHeadingNode ().fig[getActiveHeadingNode ().fig.length()-1].p[getActiveHeadingNode ().fig[getActiveHeadingNode ().fig.length()-1].p.length()-1];
	activeP['inline-graphic'][activeP['inline-graphic'].length()-1].@['mime-subtype'] = "jpg";
	activeP['inline-graphic'][activeP['inline-graphic'].length()-1].@['xlink::href'] = "CH"+pad(iChapter,3)+".MTAB"+pad(iMiscTable,3);
	activeP['inline-graphic'][activeP['inline-graphic'].length()-1].@['xlink::title'] = "Table "+iMiscTable;
	activeP['inline-graphic'][activeP['inline-graphic'].length()-1].@['xhtml::xlink'] = "http://www.w3.org/1999/xlink";
	manifest[manifest.length] = {href:"CH"+pad(iChapter,3)+"."+"MTAB"+pad(iMiscTable+1,3)};
	}