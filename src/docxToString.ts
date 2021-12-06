import fs from 'fs';
import JSZip from 'jszip';
import { parseString } from 'xml2js';

export const docxToString = async (filePath: string): Promise<string> => {
  return new Promise((resolve, reject) => {
    const docxFile = fs.readFileSync(filePath);
    // unzip the file
    const zip = new JSZip();
    zip.loadAsync(docxFile).then(function (zip) {
      // get the content of the document.xml file
      const wordFolder = zip.folder('word');
      if (!wordFolder) { reject(`An error ocurred attempting to enter to the folder 'word' of the docx file.`); return; }
      const file = wordFolder.file("document.xml");
      if (!file) { reject(`An error ocurred attempting to enter to the load the file 'document.xml' in folder 'word' of the docx file.`); return; }
      file.async('string').then(function (xMLContent) {
        parseString(xMLContent, function (err, result) {
          const paragraphs = result['w:document']['w:body'][0]['w:p'];
          let docxInTxt: string = '';

          paragraphs.forEach((paragraph: { [x: string]: { [x: string]: any[]; }[]; }) => {
            let textInTheParagraph: string = '';
            const saveTheParagraph = () => {
              textInTheParagraph += '\r\n';
              docxInTxt += textInTheParagraph;
            };
            const wRLabel = paragraph['w:r'];
            if (!wRLabel || !wRLabel.length) { saveTheParagraph(); return; }
            wRLabel.forEach((wR: { [x: string]: any[]; }) => {
              let text: string = '';
              const wTLabel = wR['w:t'];
              // check if WTLabel is an object and has the "_" property
              if (wTLabel && wTLabel.length && wTLabel[0]['_']) {
                text = wTLabel[0]['_'];
              } else {
                if (wTLabel && wTLabel.length && typeof wTLabel[0] === 'string') {
                  text = wTLabel[0];
                } else {
                  if (wTLabel && wTLabel.length && wTLabel[0]['$']) {
                    text = " ";
                  }
                }
              }
              textInTheParagraph += text;
            });
            saveTheParagraph();
          });

          resolve(docxInTxt);
        });
      });
    });
  });
};