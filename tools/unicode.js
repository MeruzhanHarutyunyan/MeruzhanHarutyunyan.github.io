var {DOMParser, XMLSerializer} = require('xmldom');
var xpath  = require("xpath");
var JsZip = require("jszip");
var fs = require("fs");
var Path = require("path");

console.log(process.argv)   
var docxInputPath = Path.join(process.cwd(), process.argv[2]); 
var strOutputPath = "output.txt";

// sometimes non-unicode font is embedded in style, and we can't detect it 
var forceArmFonts = process.argv.indexOf("--force-armenian") != -1

var fontNameMap = {
    "Courier LatArm": 1,
    'Times Armenian': 1,
    'Times LatArm': 1,

    'Russian Times': 2,
    'Times LatRus': 2,
    "Baltica": 2,
    'Courier LatRus': 2,
}

function replaceText(node) {
    if (!node.childNodes) return;
    for (var i = 0; i < node.childNodes.length; i++) {
        var child = node.childNodes[i];
        if (child.tagName === "w:t") {

            var serializer = new XMLSerializer();        
            var docx_str_new = serializer.serializeToString(node);
            var fonts = Array.from(docx_str_new.matchAll(/w:ascii="([^"]+)"/g)).map(x=>x[1])

            if (!fonts.length && forceArmFonts) {
                fonts = ['Times Armenian']
                console.log(changeEncoding(child.textContent, armMap, docx_str_new.slice(0, 100)) )
            }

            if (!fonts.length) {
                if (/[\xa0-\xff]/.test(docx_str_new))
                    console.log("no fonts", docx_str_new.slice(0, 300))
            } else if (fonts.length == 1) {
                var fontVal = fontNameMap[fonts[0]]
                if (fontVal) {
                    child.textContent = changeEncoding(child.textContent, fontVal == 1 ? armMap : rusMap, docx_str_new.slice(0, 100));
                } else {
                    if (/[\xa0-\xff]/.test(docx_str_new))
                        console.log(fonts[0], docx_str_new.slice(0, 100))
                }
            } else {
                console.log(fonts)
            }
            // child.textContent = changeEncoding(child.textContent);
            // if (/"Times Armenian"/.test(docx_str_new)) {
            //     child.textContent = changeEncoding(child.textContent);
            // }
            // else {
            //     child.textContent = "----" + changeEncoding(child.textContent) + "----";
            // }

        } else {
            replaceText(child);
        }
    }
}

// Read the docx internal xdocument
async function main() {
    var wSelect = xpath.useNamespaces({"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"});
    var docxFile = fs.readFileSync(docxInputPath);
    await JsZip.loadAsync(docxFile).then(async (zip) => {
        var docx_str = await (zip.file('word/document.xml').async("string"))
        var docx = new DOMParser().parseFromString(docx_str);

        replaceText(docx);

        var serializer = new XMLSerializer();        
        var docx_str_new = serializer.serializeToString(docx);

        var strOutputPath = docxInputPath.replace(/\.docx$/, ".unicode.docx");
        fs.writeFileSync(strOutputPath + ".xml", docx_str_new, "utf8");

        // for debugging
        fs.writeFileSync(strOutputPath + ".debug.xml", docx_str.replace(/\/[\w:]*>/g, "$&\n"), "utf8");

        zip.file('word/document.xml', docx_str_new);

        zip.generateNodeStream({type:'nodebuffer', streamFiles:true})
            .pipe(fs.createWriteStream(strOutputPath))
            .on('finish', function () {
                console.log("output`enter code here`.zip written.");
            });
    
        
        // .then(docx_str =>  
        // {

        //     var outputString = "";
        //     var paragraphElements = this.wSelect("//w:p",docx);
        //     paragraphElements.forEach(paragraphElement => 
        //     {
        //         var textElements = this.wSelect(".//w:t",paragraphElement);
        //         textElements.forEach(textElement => outputString += textElement.textContent);
        //         if(textElements.length > 0)
        //             outputString += "\n";
        //     });
        //     fs.writeFileSync(strOutputPath,outputString);
        // });
    });
}








/*************rus encodings*************************************/
var rusUni = Array(1095, 1105, 1102, 1103, 1074, 1077, 1088, 1090, 1099, 1091, 1080, 1086, 1087, 1096, 1097, 1101, 1072, 1089, 1076, 1092, 1075, 1093, 1081, 1082, 1083, 1079, 1100, 1094, 1078, 1073, 1085, 1084, 1063, 1025, 1070, 1071, 1042, 1045, 1056, 1058, 1067, 1059, 1048, 1054, 1055, 1064, 1065, 1069, 1040, 1057, 1044, 1060, 1043, 1061, 1049, 1050, 1051, 1047, 1068, 1062, 1046, 1041, 1053, 1052)
var rusb = Array(247, 184, 254, 255, 226, 229, 240, 242, 251, 243, 232, 238, 239, 248, 249, 253, 224, 241, 228, 244, 227, 245, 233, 234, 235, 231, 252, 246, 230, 225, 237, 236, 215, 168, 222, 223, 194, 197, 208, 210, 219, 211, 200, 206, 207, 216, 217, 221, 192, 209, 196, 212, 195, 213, 201, 202, 203, 199, 220, 214, 198, 193, 205, 204)
var rusph = Array(96, 45, 61, 113, 119, 101, 114, 116, 121, 117, 105, 111, 112, 91, 93, 92, 97, 115, 100, 102, 103, 104, 106, 107, 108, 122, 120, 99, 118, 98, 110, 109, 126, 95, 43, 81, 87, 69, 82, 84, 89, 85, 73, 79, 80, 123, 125, 124, 65, 83, 68, 70, 71, 72, 74, 75, 76, 90, 88, 67, 86, 66, 78, 77)
/*************armencodings**************************************/
var armUni = Array(1383, 1385, 1411, 1393, 1403, 1415, 1408, 1401, 1395, 1386, 1412, 1400, 1381, 1404, 1407, 1384, 1410, 1387, 1413, 1402, 1389, 1390, 1399, 1377, 1405, 1380, 1414, 1379, 1392, 1397, 1391, 1388, 1382, 1394, 1409, 1406, 1378, 1398, 1396, 1335, 1337, 1363, 1345, 1355, 1360, 1353, 1347, 1338, 1364, 1352, 1333, 1356, 1359, 1336, 1362, 1339, 1365, 1354, 1341, 1342, 1351, 1329, 1357, 1332, 1366, 1331, 1344, 1349, 1343, 1340, 1334, 1346, 1361, 1358, 1330, 1350, 1348, 187, 171, 1373, 96)
var armb = Array(191, 195, 247, 211, 231, 168, 241, 227, 215, 197, 249, 225, 187, 233, 239, 193, 245, 199, 251, 229, 203, 205, 223, 179, 235, 185, 253, 183, 209, 219, 207, 201, 189, 213, 243, 237, 181, 221, 217, 190, 194, 246, 210, 230, 240, 226, 214, 196, 248, 224, 186, 232, 238, 192, 244, 198, 250, 228, 202, 204, 222, 178, 234, 184, 252, 182, 208, 218, 206, 200, 188, 212, 242, 236, 180, 220, 216, 166, 167, 170, 176)
var armph = Array(49, 50, 51, 52, 53, 55, 56, 57, 48, 61, 113, 119, 101, 114, 116, 121, 117, 105, 111, 112, 91, 93, 92, 97, 115, 100, 102, 103, 104, 106, 107, 108, 122, 120, 99, 118, 98, 110, 109, 33, 64, 35, 36, 37, 42, 40, 41, 43, 81, 87, 69, 82, 84, 89, 85, 73, 79, 80, 123, 125, 124, 65, 83, 68, 70, 71, 72, 74, 75, 76, 90, 88, 67, 86, 66, 78, 77)




var armMap = {
    161: 169,// ©
    162: 167,// §
    163: 58,// :
    164: 41,// )
    165: 40,// (
    166: 187,// »
    167: 171,// «
    168: 1415,// և
    169: 46,// .
    170: 1373,// ՝
    171: 44,// ,
    172: 1418,// ֊
    174: 8230,// …
    175: 1372,// ՜
    176: 1371,// ՛
    177: 1374,// ՞
    178: 1329,// Ա
    179: 1377,// ա
    180: 1330,// Բ
    181: 1378,// բ
    182: 1331,// Գ
    183: 1379,// գ
    184: 1332,// Դ
    185: 1380,// դ
    186: 1333,// Ե
    187: 1381,// ե
    188: 1334,// Զ
    189: 1382,// զ
    190: 1335,// Է
    191: 1383,// է
    192: 1336,// Ը
    193: 1384,// ը
    194: 1337,// Թ
    195: 1385,// թ
    196: 1338,// Ժ
    197: 1386,// ժ
    198: 1339,// Ի
    199: 1387,// ի
    200: 1340,// Լ
    201: 1388,// լ
    202: 1341,// Խ
    203: 1389,// խ
    204: 1342,// Ծ
    205: 1390,// ծ
    206: 1343,// Կ
    207: 1391,// կ
    208: 1344,// Հ
    209: 1392,// հ
    210: 1345,// Ձ
    211: 1393,// ձ
    212: 1346,// Ղ
    213: 1394,// ղ
    214: 1347,// Ճ
    215: 1395,// ճ
    216: 1348,// Մ
    217: 1396,// մ
    218: 1349,// Յ
    219: 1397,// յ
    220: 1350,// Ն
    221: 1398,// ն
    222: 1351,// Շ
    223: 1399,// շ
    224: 1352,// Ո
    225: 1400,// ո
    226: 1353,// Չ
    227: 1401,// չ
    228: 1354,// Պ
    229: 1402,// պ
    230: 1355,// Ջ
    231: 1403,// ջ
    232: 1356,// Ռ
    233: 1404,// ռ
    234: 1357,// Ս
    235: 1405,// ս
    236: 1358,// Վ
    237: 1406,// վ
    238: 1359,// Տ
    239: 1407,// տ
    240: 1360,// Ր
    241: 1408,// ր
    242: 1361,// Ց
    243: 1409,// ց
    244: 1362,// Ւ
    245: 1410,// ւ
    246: 1363,// Փ
    247: 1411,// փ
    248: 1364,// Ք
    249: 1412,// ք
    250: 1365,// Օ
    251: 1413,// օ
    252: 1366,// Ֆ
    253: 1414,// ֆ
    254: 8217,// ’
    255: 8216,// ‘
}

var rusMap = {
    168: 1025,// Ё
    184: 1105,// ё
    185: 8470,// №
    190: 8230,// …
    192: 1040,// А
    193: 1041,// Б
    194: 1042,// В
    195: 1043,// Г
    196: 1044,// Д
    197: 1045,// Е
    198: 1046,// Ж
    199: 1047,// З
    200: 1048,// И
    201: 1049,// Й
    202: 1050,// К
    203: 1051,// Л
    204: 1052,// М
    205: 1053,// Н
    206: 1054,// О
    207: 1055,// П
    208: 1056,// Р
    209: 1057,// С
    210: 1058,// Т
    211: 1059,// У
    212: 1060,// Ф
    213: 1061,// Х
    214: 1062,// Ц
    215: 1063,// Ч
    216: 1064,// Ш
    217: 1065,// Щ
    218: 1066,// Ъ
    219: 1067,// Ы
    220: 1068,// Ь
    221: 1069,// Э
    222: 1070,// Ю
    223: 1071,// Я
    224: 1072,// а
    225: 1073,// б
    226: 1074,// в
    227: 1075,// г
    228: 1076,// д
    229: 1077,// е
    230: 1078,// ж
    231: 1079,// з
    232: 1080,// и
    233: 1081,// й
    234: 1082,// к
    235: 1083,// л
    236: 1084,// м
    237: 1085,// н
    238: 1086,// о
    239: 1087,// п
    240: 1088,// р
    241: 1089,// с
    242: 1090,// т
    243: 1091,// у
    244: 1092,// ф
    245: 1093,// х
    246: 1094,// ц
    247: 1095,// ч
    248: 1096,// ш
    249: 1097,// щ
    250: 1098,// ъ
    251: 1099,// ы
    252: 1100,// ь
    253: 1101,// э
    254: 1102,// ю
    255: 1103,// я
}

var ignore = {
    32: " ",
    10: "\n",
    13: "\r",
    9: "\t",
    40: "(",
    41: ")",
    42: "*",
    43: "+",
    44: ",",
    45: "-",
    46: ".",
    48: "0",
    49: "1",
    50: "2",
    51: "3",
    52: "4",
    53: "5",
    54: "6",
    55: "7",
    56: "8",
    57: "9",
    58: ":",
    59: ";",
    91: "[",
    93: "]",
    96: "`",
    8211: "–",
    8230: "…",
}
function changeEncoding(string, fontMap, errString) {
    var map = {}

    for (var i in fontMap) {
        map[i] = String.fromCharCode(fontMap[i]);
    }
    var newString = "";
    var hasErrors = false
    for (var i = 0; i < string.length; i++) {
        var char = string[i];
        var code = char.charCodeAt(0);
        if (161< code && code < 256 && !map[code] && !ignore[code]) {
            hasErrors = true
            console.log(char, code, map[code]);
        }
        var newChar = map[code] || char;
        newString += newChar;
    }
    if (hasErrors) {
        console.log(">>>>", errString)
    }
    return newString;
}



main();



function renderFontPage() {

result=`<style>
.arm {
    font-family: "Times Armenian";
}

table {
  border-collapse: collapse;
  border: 2px solid rgb(140 140 140);
  font-family: sans-serif;
  font-size: 0.8rem;
  varter-spacing: 1px;
}

caption {
  caption-side: bottom;
  padding: 10px;
  font-weight: bold;
}

thead,
tfoot {
  background-color: rgb(228 240 245);
}

th,
td {
  border: 1px solid rgb(160 160 160);
  padding: 8px 10px;
}

td:last-of-type {
  text-align: center;
}

tbody > tr:nth-of-type(even) {
  background-color: rgb(237 238 242);
}

tfoot th {
  text-align: right;
}

tfoot td {
  font-weight: bold;
}

</style>
<table>`
for (var i =0 ;i < 0xff + 10; i++) {
    var ch = String.fromCharCode(i)
    var uch = map[i] || ''
    result += `<tr><td>${i}</td>, <td class="arm">${ch}</td>  <td class="uni">${ch}</td>
    <td contentEditable>${uch}</td>
    </tr>`
}
result+= "</table>"
document.body.innerHTML = result; 0


    function read() {
        var map = {}
        var str = "{\n"
        document.querySelectorAll("tr").forEach(x => {
            if (x.lastElementChild.textContent.trim()) {
            var from = x.firstElementChild.textContent
            var to = x.lastElementChild.textContent.trim()
            map[from] = to
            str += "    " + from + ": " + to.charCodeAt(0) + ",// "+ to + "\n"
            }
        })
        console.log(map)
        return str+"}"
    }
}