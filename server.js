



// Use library in Nodejs:     "jszip", "pptxgenjs", "xml2js"

const fs = require('fs');
const JSZip = require('jszip');
const xml2js = require('xml2js');

const processPptx = async (filePath, outputFilePath) =>   {
    try {
        // fs.readFileSync: Read the entire contents of a PPTX file.
        const fileData = fs.readFileSync(filePath);

        const zip = await JSZip.loadAsync(fileData);

        // zip.files: List of all files inside the zip file.
        // .filter: Only get files with paths starting with ppt/slides/slide (files containing slide content).
        const slideFiles = Object.keys(zip.files).filter((fileName) =>
            fileName.startsWith('ppt/slides/slide')
        );
        console.log(slideFiles);
        

        // zip.files[slideFile].async('string'): Reads the contents of the slide file as a string.
        // xml2js.parseStringPromise(content): Parses an XML string into a JavaScript object.
        for (const slideFile of slideFiles) {
            const content = await zip.files[slideFile].async('string');
            // console.log(content);
            const parsedContent = await xml2js.parseStringPromise(content);
            

            // parsedContent['p:sld']['p:cSld'][0]['p:spTree'][0]['p:sp']: Get a list of shapes in the slide.
            // shape['p:txBody']: Check if the shape has text content.
            // paragraphs: Iterate through the paragraphs in the text content.
            // p['a:r']: Iterate through the child elements containing text.
            // r['a:t']: Get the text value and convert it to uppercase using .toUpperCase().

            const shapes = parsedContent['p:sld']['p:cSld'][0]['p:spTree'][0]['p:sp'];
            if (shapes) {
                shapes.forEach((shape) => {
                    if (shape['p:txBody']) {
                        const paragraphs = shape['p:txBody'][0]['a:p'];
                        paragraphs.forEach((p) => {
                            if (p['a:r']) {
                                p['a:r'].forEach((r) => {
                                    if (r['a:t']) {
                                        r['a:t'][0] = r['a:t'][0].toUpperCase(); // Chuyển sang IN HOA
                                    }
                                });
                            }
                        });
                    }
                });
            }
            // xml2js.Builder: Creates an object to convert data from JavaScript back to XML.
            // zip.file(slideFile, updatedXml): Updates the edited content to the ZIP file.            

            const builder = new xml2js.Builder();
            const updatedXml = builder.buildObject(parsedContent);

            zip.file(slideFile, updatedXml);
        }

        // zip.generateAsync: Tạo file nén mới từ nội dung đã chỉnh sửa.
        // fs.writeFileSync: Ghi file nén mới dưới dạng file PPTX.

        const updatedData = await zip.generateAsync({ type: 'nodebuffer' });
        fs.writeFileSync(outputFilePath, updatedData);
        console.log(`File PPTX create successfull: ${outputFilePath}`);
    } catch (err) {
        console.error('Error hanlde file PPTX:', err);
    }
}


const inputFilePath = 'data/input/ThinkPrompt_BE_testing.pptx';
const outputFilePath = 'data/output/ConvertUppercase_ThinkPrompt_BE_testing.pptx';

processPptx(inputFilePath, outputFilePath);
