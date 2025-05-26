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
    console.log("Processing " + path);
    return code.replace(/class="book-card">[\s\S]*?<\/div>/g, function(match) {
        match = match.replace(/ *<a href="(.*?)"[^>]*>(PDF|EPUB)<\/a>\n?/g, "")
        console.log("Removing old links from " + match);
        return match.replace(/( *)<a href="(.*?).html"[^>]*>(Կարդալ|HTML)<\/a>\n/g, function(all, indent, href) {
            return `${indent}<a href="${href}.html">HTML</a>\n` +
                `${indent}<a href="${href}.pdf">PDF</a>\n` +
                `${indent}<a href="${href}.epub">EPUB</a>\n`;
        });
    })
})