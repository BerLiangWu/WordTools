# WordTools
UserconfigcontentReplaceTheFileKeyWords by python  
## How Use  
The ProjectFile/EXE/dist is EXE file Run ON X64Windows  
1.The file RootPath Configfilefolder Generated IF you First Use  
2.ProjectFile/Config/userconfig.json  
{  
    "ReplaceCount": "2",   //How many content Will replace  
    "[A]": "You DIY content",  You DIY content replace [A]  
    "[B]":"You DIY content"   You DIY content replace [A]  
 }  
    matters need attention matter：  
    You have to label the document you need to replace with keywords yourself，KeyWords:[A] [B] [C] now support to [H]  
    for example  ：
    You have one documents that need to be replaced automatically  
    name:  
    age:  
    So You Should be The following configuration files  
    {  
    "ReplaceCount": "2",   //How many content Will replace  
    "[A]": "AntiLiang",  You DIY content replace [A]  
    "[B]":"114514"   You DIY content replace [A]  
 }   
 So this tool is actually a convenient tool for you to fill in different documents with repeated content  
 This tool supports multiple documents, just put multiple documents in the same directory  
 ## 怎么用？
 工程目录下有一个EXE目录，里面的dist里放着已经打包成WIN64环境可直接使用的文件 
 1.当你第一次使用时，将会自动生成配置目录config 里面有个名为userconfig.json  的文件  
 {  
    "ReplaceCount": "2",   //你有多少个内容需要被替换  
    "[A]": "You DIY content",  //你标记的关键字A将会被替换上你指定的内容 
    "[B]":"You DIY content"   //你标记的关键字B将会被替换上你指定的内容 
 }  
 特别事项：  
 本工具需要你去每个文件里面标上关键字[A],[B],[C]...等 目前支持到[H]
 例子：
   你有一份文件，你需要填入名字和年龄，并且日后可能作为模板使用，你可能会多次填入所以你就得这样配置json文件
      {  
   "ReplaceCount": "2",   //你有多少个内容需要被替换  
    "[A]": "AntiLiang",  //你标记的关键字A将会被替换上你指定的内容 
    "[B]":"114514"   //你标记的关键字B将会被替换上你指定的内容 
 }  
 当程序运行时他会遍历目录下每个word文件，并且识别里面的关键标记，将标记替换成jsonkey里面的内容，所以你不难发现  
 这个是方便你多个重复内容要写入不同文档的便捷工具，一次配置，日后模板文件就能便捷替换，本工具是支持多文件替换的，只需要将这些文件放在同一个目录下即可  
 
