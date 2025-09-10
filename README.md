# RescueGamma

解决 https://gamma.app/ 导出PPT的字体异常和渐变效果异常的专用小工具

# 项目背景

<img width="700" height="750" alt="image" src="https://github.com/user-attachments/assets/2b522b8e-8665-44a8-a333-348e424651a5" />

目前比较流行且好用的一款AI做PPT的应用——Gamma，在导出PPT到本地时，字体和渐变效果等会显示异常（官网已经明确提醒就是会异常，且实测安装它提供的字体包也无法解决问题），我平时用Gamma次数比较多，调字体和配色这种工作太繁琐费时了，因此我决定开发一款软件，解决这个问题。

# 软件核心库

主要用到了Python-pptx和xml的相关库，界面用PySide6实现

# 主要功能

1. 根据字体配置功能实现批量的字体字号替换（同时处理语言类型问题修复，解决PPT语句审查红色波浪线问题）；
   
   <img width="535" height="273" alt="image" src="https://github.com/user-attachments/assets/15c2d490-fe1c-4a10-9bd5-4d9a98da6b12" />
   
3. 复刻PPT的文字渐变色设置，实现对目标渐变色精确取色，生成渐变色配置
   
   <img width="798" height="774" alt="image" src="https://github.com/user-attachments/assets/4d9027da-3524-4e53-9ba9-61ccc8e3533d" />

   <img width="543" height="242" alt="image" src="https://github.com/user-attachments/assets/0a62b0d0-104e-428f-aabe-651ee1464708" />
 
5. 根据渐变色配置，对指定的字号、字体完成批量渐变色配置应用
   
   <img width="442" height="626" alt="image" src="https://github.com/user-attachments/assets/b81aa1e4-0ba2-4f33-83c4-263735ce94df" />

7. 调色板，点击复制单个色号
   
   <img width="396" height="445" alt="image" src="https://github.com/user-attachments/assets/35d0aeb2-f21b-4c4f-bb80-6f7458558bb2" />

9. 深色主题切换支持
    
    <img width="796" height="725" alt="image" src="https://github.com/user-attachments/assets/a53d1ce6-d1d6-482c-aa5f-9a22ffab29c1" />

    <img width="796" height="729" alt="image" src="https://github.com/user-attachments/assets/e767f9d3-909f-4a01-a190-861aad755e00" />
# 打包分发

我发布了用nuitka打包的免安装版，Windows10、11应该可以正常使用

# 一杯咖啡

大家觉得这个软件有用、能帮到自己的话，可以打赏我一杯瑞幸香草丝绒咖啡喝~

- 微信号：Moss_Go
- 邮箱：2025aibear@gmail.com / edu_datacenter@163.com
- 一杯咖啡：
  
<img width="228" height="225" alt="image" src="https://github.com/user-attachments/assets/1fc41be3-98b6-402d-98ea-4103d98b9026" />

    




