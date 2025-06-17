var {DOMParser, XMLSerializer} = require('xmldom');
var xpath  = require("xpath");
var JsZip = require("jszip");
var fs = require("fs");
var Path = require("path");

console.log(process.argv)   
var docxInputPath1 = Path.join(process.cwd(), process.argv[2]); 
var docxInputPath2 = Path.join(process.cwd(), process.argv[3]); 
var strOutputPath = "output.txt";

 
// Read the docx internal xdocument
async function extract(docxInputPath) {
    var docxFile = fs.readFileSync(docxInputPath);
    await JsZip.loadAsync(docxFile).then(async (zip) => {
        var docx_str = await (zip.file('word/document.xml').async("string"))
        
        var docx_str_new = docx_str
            .replace(/<w:softHyphen\/>/g, "")
            .replace(/<\/w:t><w:t>/g, "")
            .replace(/<w:t[^>]*>/g, "\n<w:t>")

        fs.writeFileSync(docxInputPath + ".xml", docx_str_new, "utf8");
    });
}
async function main() {
    await extract(docxInputPath1)
    await extract(docxInputPath2)
}

main();

