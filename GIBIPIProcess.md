# 目录

* [一. 原始数据](#一--原始数据)  
* [二. 宏ReadGIBIPI导入excel并将GIBIPI分别存放在三张表中](#二--宏ReadGIBIPI导入excel并将GIBIPI分别存放在三张表中)   
* [三. SAS操作过程与结果](#三--SAS操作过程与结果)   
    * [1. 录入原始数据](#1--录入原始数据)   
    * [2. 两因素方差分析 GLM](#2--两因素方差分析-GLM)  
    * [3. 检验效能计算](#3--检验效能计算)  
    * [4. 样本量估算](#4--样本量估算)  
* [四. PASS操作过程与结果](#四--PASS操作过程与结果)  
    * [1. 检验效能计算](#1--检验效能计算)  
    * [2. 给对gender因素的检验估算样本量](#2--给对gender因素的检验估算样本量)  
    * [3. 给对education因素的检验估算样本量](#3--给对education因素的检验估算样本量)   
* [五. Reference](#五--Reference)  
    
&ensp;&ensp;&ensp;&ensp;  


# 一  原始数据  
&ensp;&ensp;&ensp;&ensp;牙龈炎临床试验中GIBIPI指标原始excel数据如下：Origin.xlsx  
![image](https://github.com/TracyHuo/GIBIPIProcess/blob/master/Image/1.PNG);  
&ensp;&ensp;&ensp;&ensp;   
&ensp;&ensp;&ensp;&ensp;   
# 二  宏ReadGIBIPI导入excel并将GIBIPI分别存放在三张表中  
* **代码**：   
```
/****************************************************************************
宏ReadGIBIPI的作用是读取外部excel数据集，然后将数据分隔成GI,BI,PI三个数据集。
*****************************************************************************/
%MACRO ReadGIBIPI ;

  /********************************************************************************
   1. 导入原始数据，由filepath指定路径，sheetname指定表明，导入到mylib库的Origin数据集
   *********************************************************************************/
    
    PROC IMPORT DATAFILE = "F:\Clinical trials\GIBIPI\Origin.xlsx"  
	/*导入excel时，excel文件应处于关闭状态，否则报错*/
            OUT = mylib.Origin
            DBMS = xlsx
            REPLACE;
            RANGE = "Base$A1:AR90"; /*可以再定义一个宏参数设置受试者个数，以改变90的固定值*/
            GETNAMES = NO ;
    RUN;

  /*****************************************************************************************
	2. 将Origin中的GI, BI, PI数据提取出来，分别存放到mylib.GI, mylib.BI, mylib.PI三个数据集中
   *****************************************************************************************/

    DATA mylib.GI(keep = Name ID B--V  
                  rename=(B=C_1 C=C_2 D=C_3 E=C_4 F=C_5 G=C_6 H=C_7 I=C_8 J=C_9 K=C_10 L=C_11 
                          M=C_12 N=C_13 O=C_14 P=C_15 Q=C_16 R=C_17 S=C_18 T=C_19 U=C_20 V=C_21))  
         mylib.BI(keep = Name ID B--V 
                  rename=(B=C_1 C=C_2 D=C_3 E=C_4 F=C_5 G=C_6 H=C_7 I=C_8 J=C_9 K=C_10 L=C_11 
                          M=C_12 N=C_13 O=C_14 P=C_15 Q=C_16 R=C_17 S=C_18 T=C_19 U=C_20 V=C_21))    
         mylib.PI(keep = Name ID X--AR 
                  rename=(X=C_1 Y=C_2 Z=C_3 AA=C_4 AB=C_5 AC=C_6 AD=C_7 AE=C_8 AF=C_9 AG=C_10 
                          AH=C_11 AI=C_12 AJ=C_13 AK=C_14 AL=C_15 AM=C_16 AN=C_17 AO=C_18 
                          AP=C_19 AQ=C_20 AR=C_21))  ;
         /*这里改列名的方法很繁琐。但rename=(B--V = C_1-C_21)报错：RENAME 不支持指定变量列表的表单。
		选项被忽略。有比较好的改列名方法吗？*/

         LENGTH Name $9 ID $ 2; /*大小需查看mylib.Origin里由SAS导入数据时自动确定的Name列的长度*/
         RETAIN Name ID;
 
		 /*一位受试者在原始文件中占30行，以下通过SET--IF--OUTPUT的方式把相应的行和列输出到GI/BI/PI表。
		   用linenum = mod(_N_, 30);的方式实现类似循环的效果。但多分支的IF--ELSE仍显繁琐。
		   有无更简单的方法？*/
         SET mylib.Origin;
         linenum = mod(_N_, 30);  
         IF  linenum=1 THEN DO; 
                               Name=C; ID=H; /*Name和ID可以retain*/ 
                               OUTPUT; /*表示DATA语句后的GI,BI,PI三个数据集都输出*/
                            END;   

         ELSE IF  linenum=4 THEN DO; OUTPUT mylib.GI mylib.PI;END;
         /*可以用B--V么？另外，只想把第2到22列输出给GI数据集，在OUTPUT后的数据集后使用数据集选项可以吗？*/
         ELSE IF linenum=5 THEN DO; OUTPUT mylib.BI mylib.PI; END; 
         ELSE IF linenum=6 THEN DO; OUTPUT mylib.GI;  END;
         ELSE IF linenum=7 THEN DO; OUTPUT mylib.BI;  END;             
         ELSE IF linenum IN (9,10) THEN DO; OUTPUT mylib.PI;  END; 
         ELSE IF linenum=11 THEN DO; OUTPUT mylib.GI;  END;           
         ELSE IF linenum=12 THEN DO; OUTPUT mylib.BI;  END;                
         ELSE IF linenum=13 THEN DO; OUTPUT mylib.GI;  END;        
         ELSE IF linenum=14 THEN DO; OUTPUT mylib.BI mylib.PI;END;  
         ELSE IF linenum=15 THEN DO; OUTPUT mylib.PI;  END;
         ELSE IF linenum=18 THEN DO; OUTPUT mylib.GI;  END;
         ELSE IF linenum=19 THEN DO; OUTPUT mylib.BI mylib.PI;   END;
         ELSE IF linenum=20 THEN DO; OUTPUT mylib.GI mylib.PI;END;  
         ELSE IF linenum=21 THEN DO; OUTPUT mylib.BI;  END;   
         ELSE IF linenum IN (25,27) THEN DO; OUTPUT mylib.GI;  END;
         ELSE IF linenum IN (26,28) THEN DO; OUTPUT mylib.BI;  END;

    RUN;

	/*******************************************************************************
	  3.删除以上所得GI,BI,PI表里每位受试者的首行（每位受试者占9行），此行是为了retain 
	    Name和ID变量而产生的，是多余的，需要删除。直接对GI,BI,PI原表修改即可。
	    此Pro_Delete实现删除每位受试者的首行的目的，这里写成了宏ReadGIBIPI内的嵌套宏。
	 *******************************************************************************/
	
	%MACRO Pro_Delete(dataset= );
	    DATA &dataset(drop=n) ;
		    n=nobs/9 ;
			SET &dataset nobs=nobs;
			IF mod(_N_, 9) = 1 THEN DELETE;
		RUN;
	%MEND Pro_Delete;

    /*******************************************************************************
	  4.对GI,BI,PI三张表执行Pro_Delete，删除每位受试者的首行。这样，每位受试者占8行。
	 *******************************************************************************/
	
	%Pro_Delete(dataset=mylib.GI)
	%Pro_Delete(dataset=mylib.BI)
	%Pro_Delete(dataset=mylib.PI)


%MEND ReadGIBIPI ;

%ReadGIBIPI ;
```

* **结果**：  

