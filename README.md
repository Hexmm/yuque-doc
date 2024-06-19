# yuque-docx
语雀文档转docx

## 准备

- 安装[pandoc](https://www.pandoc.org/), 转换md到docx的工具
- 安装[yuque-dl](https://www.npmjs.com/package/yuque-dl),爬取语雀文档内容

## 开始
1. yuque-dl 'url' 下载
2. 获取下载md文件在所在目录的绝对路径，举例为 root/**/download/文档名称
3. 替换 #2 的路径到 convert.cjs 中的 mdDirPath值
4. 执行启动命令
    ```
     yarn start
    ```
   
## 注意
### 已知问题
1. 如果语雀文档存在json代码块，json中存在 '$someKey' 的情况，pandoc会识别为表达式，导致报错。 
可将md中$修改为\\$,在最终的docx文件中再自行恢复
