var fs = require("fs"); 

function replaceInFiles(path, fn) {
    if (Array.isArray(path)) {
        path.forEach(function(x) {
            replaceInFiles(x, fn);
        });
        return;
    }
    var stat = fs.lstatSync(path);
    if (stat.isDirectory()) {
        var files = fs.readdirSync(path);
        files.forEach(function (x) {
            if (x == "node_modules" || x == ".git") return;
            replaceInFiles(path + "/" + x, fn);
        });
    } else if (stat.isFile()) {
        var text = fs.readFileSync(path, "utf8");
        var newText = fn(text, path);
        if (newText != text && typeof newText == "string") {
            console.log(path);
            fs.writeFileSync(path, newText, "utf8");
        }
    }
}


replaceInFiles(__dirname + "/../index.html", function(code, path) {

    var files = fs.readdirSync(__dirname + "/../books")
        .filter(function(x) {
            return /\.docx$/.test(x) && !/~/.test(x)
        }).map(function(x) {
            return x.slice(0, -5)
        })
    console.log(files) 
    var links = Array.from(
        code.matchAll(/books\/(.*)\.html/g)
    ).map(x=>x[1])

    files.forEach(function(x) {
        if (links.includes(x)) return
        console.log(x)
        var i = code.lastIndexOf('<div class="books-grid">')
        var startRe = /^(\s*)<div class="book-card">/gm
        startRe.lastIndex = i
        var m = startRe.exec(code)
        var startIndex = m.index;
        var endStr = "\n" + m[1] + "</div>"

        var endIndex = code.indexOf(endStr, m.index)
        console.log(startIndex, endStr, endIndex)
        if (endIndex == -1) 
            throw new Error("endindex not found")
        var str = code.slice(startIndex, endIndex + endStr.length)
        str = str.replace(/"books\/.*?\.(?:html|pdf|epub)"/g, '"books/' + x + '.$1"')
        
        code = code.slice(0, startIndex) + str + "\n" + code.slice(startIndex, -1)
    })

    console.log("Processing " + path);
    return code.replace(/class="book-card">[\s\S]*?<\/div>/g, function(match) {
        match = match.replace(/ *<a href="(.*?)"[^>]*>(PDF|EPUB)<\/a>\n?/g, "")
        // console.log("Removing old links from " + match);
        return match.replace(/( *)<a href="(.*?).html"[^>]*>(Կարդալ|HTML)<\/a>\n/g, function(all, indent, href) {
            return `${indent}<a href="${href}.html">HTML</a>\n` +
                `${indent}<a href="${href}.pdf">PDF</a>\n` +
                `${indent}<a href="${href}.epub">EPUB</a>\n`;
        });
    })
})