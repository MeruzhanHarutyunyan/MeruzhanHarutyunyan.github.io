const mammoth = require('mammoth');
const cheerio = require('cheerio');
const fs = require('fs-extra');
const path = require('path');
const sanitizeHtml = require('sanitize-html');
const Epub = require('epub-gen'); 

class BookConverter {
    constructor(inputFile, outputDir) {
        this.inputFile = inputFile;
        this.outputDir = outputDir;
        this.bookTitle = path.basename(inputFile, '.docx');
        this.htmlContent = '';
        this.toc = [];
    }

    async convert() {
        try {
            // Create output directory if it doesn't exist
            await fs.ensureDir(this.outputDir);

            // Convert Word to HTML
            await this.convertToHtml();

            // Generate EPUB
            // await this.generateEpub();

            console.log('Conversion completed successfully!');
        } catch (error) {
            console.error('Error during conversion:', error);
        }
    }

    async convertToHtml() {
        // Convert Word to HTML using mammoth
        const result = await mammoth.convertToHtml({ path: this.inputFile });
        this.htmlContent = result.value;

        // Process the HTML content
        const $ = cheerio.load(this.htmlContent);

        // Extract and create table of contents
        this.createTableOfContents($);

        // Clean and format the HTML
        this.cleanAndFormatHtml($);

        // Save the HTML file
        const htmlOutput = path.join(this.outputDir, `${this.bookTitle}.html`);
        await fs.writeFile(
            htmlOutput,
            $.html()
                .replace(/<\/(p|h\d?)>/g, "$&\n")
                .replace(/\xad/g, "")
        );
        console.log(`HTML file created: ${htmlOutput}`);
    }

    createTableOfContents($) {
        const toc = [];
        let tocHtml = `<style>
        
header {
    background-color: #fff;
    box-shadow: 0 2px 5px rgba(0, 0, 0, 0.1);
    position: fixed;
    width: 100%;
    top: 0; left: 0;
    z-index: 1000;
}
nav {
    max-width: 1200px;
    margin: 0 auto;
    padding: 1rem;
    display: flex;
    justify-content: space-between;
    align-items: center;
}
        </style><header>
        <nav>
            <div class="logo"><a href="../index.html">Մերուժան Հարությունյան</a></div>
        </nav>
    </header>` + '<div class="toc">\n<h2></h2>\n<ul>\n';

        // Find all headings (h1, h2, h3)
        $('h1, h2, h3').each((index, element) => {
            const $element = $(element);
            const level = parseInt(element.name[1]);
            const text = $element.text();
            const id = `section-${index}`;

            // Add ID to the heading
            $element.attr('id', id);

            // Add to TOC array
            toc.push({
                level,
                text,
                id
            });

            // Add to TOC HTML
            tocHtml += `<li class="toc-level-${level}"><a href="#${id}">${text}</a></li>\n`;
        });

        tocHtml += '</ul>\n</div>';
        this.toc = toc;

        // Insert TOC at the beginning of the content
        $('body').prepend(tocHtml);
    }

    cleanAndFormatHtml($) {
        // Add basic styling
        $('head').append(`
            <meta charset="utf-8" />
            <style>
                body { font-family: Arial, sans-serif; line-height: 1.6; max-width: 800px; margin: 0 auto; padding: 20px; }
                .toc { margin-bottom: 30px; padding: 20px; background: #f5f5f5; border-radius: 5px; }
                .toc ul { list-style-type: none; padding-left: 20px; }
                .toc-level-1 { font-weight: bold; }
                .toc-level-2 { padding-left: 20px; }
                .toc-level-3 { padding-left: 40px; }
                h1, h2, h3 { color: #333; }
                img { max-width: 100%; height: auto; }
            </style>
        `);

        // Sanitize HTML
        const sanitizedHtml = sanitizeHtml($.html(), {
            allowedTags: ['h1', 'h2', 'h3', 'p', 'a', 'img', 'ul', 'li', 'div', 'span', 'br', 'strong', 'em'],
            allowedAttributes: {
                'a': ['href', 'id'],
                'img': ['src', 'alt'],
                '*': ['class', 'id']
            }
        });

        // Update the content with sanitized HTML
        this.htmlContent = sanitizedHtml;
    }

    // async generateEpub() {
    //     const options = {
    //         title: this.bookTitle,
    //         author: 'Author Name', // You might want to extract this from the document
    //         publisher: 'Publisher Name',
    //         cover: '', // Path to cover image if available
    //         content: this.toc.map(item => ({
    //             title: item.text,
    //             data: `<h${item.level}>${item.text}</h${item.level}>`
    //         }))
    //     };

    //     const epubOutput = path.join(this.outputDir, `${this.bookTitle}.epub`);
    //     const epub = new Epub(options, epubOutput);

    //     await new Promise((resolve, reject) => {
    //         epub.promise.then(() => {
    //             console.log(`EPUB file created: ${epubOutput}`);
    //             resolve();
    //         }).catch(reject);
    //     });
    // }
}

async function processBooksDirectory() {
    const booksDir = path.join(__dirname, '..', 'books');
    
    try {
        // Ensure books directory exists
        await fs.ensureDir(booksDir);
        
        // Get all subdirectories
        const files = await fs.readdir(booksDir);

        // Find docx files
        const docxFiles = files.filter(file => {
            return file[0] !== '~' &&
                file.toLowerCase().endsWith('.docx') 
        });

        for (const docxFile of docxFiles) {
            await processBook(docxFile, booksDir, files);
        }
        console.log('Finished processing all book directories');
    } catch (error) {
        console.error('Error processing books directory:', error);
    }
}

async function processBook(docxFile, dirPath, files) {
    try {
        const docxPath = path.join(dirPath, docxFile);
        const baseName = path.basename(docxFile, '.docx');
        
        // Check if corresponding HTML or EPUB files exist
        const htmlExists = files.includes(`${baseName}.html`);
        // const epubExists = files.includes(`${baseName}.epub`);
        
        console.log(`Processing ${docxFile} in ${path.basename(dirPath)}`);

        const converter = new BookConverter(docxPath, dirPath);
        await converter.convert();


        // If either HTML or EPUB is missing, process the document
        // if (!htmlExists) {
            
            
            
        //     console.log(`Completed processing ${docxFile}`);
        // } else {
        //     console.log(`Skipping ${docxFile} - both HTML and EPUB files already exist`);
        // } 
    } catch (error) {
        console.error(`Error processing directory ${dirPath}:`, error);
    }
}


processBooksDirectory(); 