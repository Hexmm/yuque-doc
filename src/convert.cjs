const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');
const DocxMerger = require('docx-merger');

// TODO 之前之前，需要配置 yuqueq-dl下载的md所在的目录的绝对路径，举例为 root/**/download/xx
const mdDirPath = ''

const referencePath = path.join(__dirname, 'custom-reference.docx');

// 读取 index.md 并提取 Markdown 文件列表
const indexPath = path.join(mdDirPath, 'index.md');
const indexContent = fs.readFileSync(indexPath, 'utf8');

let currentLevel = 0;
const markdownFilesWithLevels = indexContent.match(/\[.*?\]\((.*?)\)|(#+\s.*)/g).map(link => {
    const matches = link.match(/\[.*?\]\((.*?)\)/);
    if (matches) {
        return {
            file: matches[1],
            parentLevel: currentLevel
        }
    } else {
        currentLevel = (link.match(/#/g) || []).length
        return {
            file: null,
            content: link,
            selfLevel: currentLevel
        }
    }
})

// 构建输出文件夹路径
const outputDirectory = path.join(__dirname, 'output');
if (!fs.existsSync(outputDirectory)) {
    fs.mkdirSync(outputDirectory, { recursive: true });
}

outputPathList = [];

const getRelativePath = (path) => {
    const index = path.indexOf("/src/");
    let result = '';
    if (index !== -1) {
        // 获取 /src/ 后的子字符串
        result = path.substring(index + 4); // 5 是 "/src/" 的长度
        console.log(result);
    } else {
        console.log("路径中未找到 /src/");
    }
    return result;
}

// 遍历每个 Markdown 文件，分别转换为 Word 文档
let titleLines = [];
markdownFilesWithLevels
    // .filter((item, index) => index < 50)
    .forEach((item, index) => {
    const {
        file: markdownFile,
        parentLevel,
        content
    } = item;
    if (!markdownFile) { // 如果没有对应的md文件，即为标题，为标题生成一个docx文件
        titleLines.push(content + '\n\n');
        const mdDirectory = outputDirectory;
        const outputFileName = 'title-' + index + content.replace(/#|\s/g, '') + '.md';
        const outputFilePathMd = path.join(mdDirectory,'/titlesmd/', outputFileName);
        const outputFilePathDoc = path.join(mdDirectory,'/titlesdoc/', path.basename(outputFileName, '.md')) + '.docx';
        outputPathList.push(outputFilePathDoc);
        // 创建输出目录（如果不存在）
        if (!fs.existsSync(path.dirname(outputFilePathMd))) {
            fs.mkdirSync(path.dirname(outputFilePathMd), { recursive: true });
        }
        if (!fs.existsSync(path.dirname(outputFilePathDoc))) {
            fs.mkdirSync(path.dirname(outputFilePathDoc), { recursive: true });
        }
        fs.writeFileSync(outputFilePathMd, content, 'utf8');
        // 拼接 Pandoc 命令
        const pandocCommand = `pandoc "${outputFilePathMd}" --reference-doc ${referencePath} -o "${outputFilePathDoc}"`;

        try {
            // 执行 Pandoc 命令，切换工作目录到 md 文件所在目录
            execSync(pandocCommand, { cwd: mdDirectory });
            console.log('转换完成：',index, outputFileName);
        } catch (error) {
            console.error('转换失败：', error);
        } finally {
            // 删除临时文件
            fs.unlinkSync(outputFilePathMd);
        }
        return;
    }
    const inputFilePath = path.join(mdDirPath, markdownFile);
    const mdDirectory = path.dirname(inputFilePath);
    const outputFileName = path.basename(markdownFile, '.md') + '.docx';
    const childDocxDirectory = path.join(outputDirectory,getRelativePath(mdDirectory))
    const outputFilePath = path.join(childDocxDirectory, outputFileName);
    outputPathList.push(outputFilePath);

    // 创建输出目录（如果不存在）
    if (!fs.existsSync(childDocxDirectory)) {
        fs.mkdirSync(childDocxDirectory, { recursive: true });
    }

    // 读取并修改 Markdown 文件内容
    let mdContent = fs.readFileSync(inputFilePath, 'utf8');
    mdContent = mdContent.split('\n').map(line => {
        let match = line.match(/^(#+)\s/); // 匹配标题行
        if (match) {
            return '#'.repeat(parentLevel) + line;
        }
        return line;
    }).join('\n');

    mdContent = titleLines.join('') + mdContent;
    titleLines = [];
    // 保存修改后的内容到临时文件
    const tempFilePath = path.join(mdDirectory, `temp-${path.basename(markdownFile)}`);
    fs.writeFileSync(tempFilePath, mdContent, 'utf8');

    // 拼接 Pandoc 命令
    const pandocCommand = `pandoc "${path.basename(tempFilePath)}" --reference-doc ${referencePath} -o "${outputFilePath}"`;

    try {
        // 执行 Pandoc 命令，切换工作目录到 md 文件所在目录
        execSync(pandocCommand, { cwd: mdDirectory });
        console.log('转换完成：',index, outputFileName);
    } catch (error) {
        console.error('转换失败：', error);
    } finally {
        // 删除临时文件
        fs.unlinkSync(tempFilePath);
    }
});

console.log('所有文件转换完成。');

fsList = outputPathList.map(pathStr => {
    return fs
        .readFileSync(pathStr, 'binary');
})

console.log('合并word文件中……');
const docx = new DocxMerger({},fsList);
docx.save('nodebuffer',function (data) {
    fs.writeFile(path.join(__dirname, 'output/output.docx'), data, function(err){/*...*/});
});
console.log('处理完毕，结果文件：',path.join(__dirname, 'output/output.docx'));
