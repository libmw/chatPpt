import Presentation from './Presentation.js';

const replaceXmlContent = (XmlContent, placeHolder, targetData) => {
  let replacedContent;
  const replaceReg = new RegExp(`<a:p>(?:.(?!<\\/a:p>))+${placeHolder}.+?<\/a:p>`);
  const placeHolderReg = new RegExp(placeHolder);
  if (typeof targetData === 'string') {
    replacedContent = XmlContent.replace(placeHolderReg, targetData);
  } else if (Array.isArray(targetData)) {
    replacedContent = XmlContent.replace(replaceReg, (match) => {
      return targetData
        .map((dataItem) => {
          return match.replace(placeHolderReg, dataItem);
        })
        .join('');
    });
  } else if (Object.keys(targetData).length > 0) {
    replacedContent = XmlContent.replace(replaceReg, (match) => {
      return Object.keys(targetData)
        .map((key) => {
          return match.replace(placeHolderReg, `${key}: ${targetData[key]}`);
        })
        .join('');
    });
  }

  return replacedContent;
};

export default async function generateByText(templatePath, outPath, data) {

  const myPresentation = new Presentation();

  await myPresentation.loadFile(templatePath);

  const newSlides = [];

  const cloneSlide1 = myPresentation.getSlide(1).clone();
  cloneSlide1.content = cloneSlide1.content.replace(/\[Title\]/g, '我的PPT');
  newSlides.push(cloneSlide1);

  const subTitleListSlide = myPresentation.getSlide(2).clone();
  const subTitles = Object.keys(data);
  subTitleListSlide.content = replaceXmlContent(
    subTitleListSlide.content,
    '\\[SubTitle\\]',
    subTitles
  );
  newSlides.push(subTitleListSlide);

  subTitles.forEach((subTitle, index) => {
    const subTitleIndexSlide = myPresentation.getSlide(3).clone();
    const subTitleContentSlide = myPresentation.getSlide(4).clone();

    subTitleIndexSlide.content = subTitleIndexSlide.content
      .replace('[SubTitleIndex]', index + 1)
      .replace('[SubTitle]', subTitle);

    subTitleContentSlide.content = replaceXmlContent(
      subTitleContentSlide.content,
      '\\[SubContent\\]',
      data[subTitle]
    );
    subTitleContentSlide.content = subTitleContentSlide.content.replace('[SubTitle]', subTitle);

    newSlides.push(subTitleIndexSlide);
    newSlides.push(subTitleContentSlide);
  });

  const newPresentation = await myPresentation.generate(newSlides);

  await newPresentation.saveAs(outPath);

  console.log('out put successfully! ');
}
