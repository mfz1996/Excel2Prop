直接运行方式：
    1.点击run-Edit configuration-在program argument里填写excel路径-apply-ok
    2.运行后会生成generate.properties
    
jar包运行方式：
    1.在Excel2Prop_jar文件夹里有jar包，直接运行java -jar Excel2Prop.jar test.xlsx
    2.即可看到在当前目录下生成properties文件
    
打jar包流程：
    参考https://blog.csdn.net/Thousa_Ho/article/details/72799871
    如果新生成的jar包提示没有主类，可以在Excel2Prop.jar\META-INF\MANIFEST.MF中加入Main-Class: com.huawei.demo.Main