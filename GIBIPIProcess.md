# 目录

* [一. 原始数据](#一--原始数据)  
* [二. 宏ReadGIBIPI导入excel并将GIBIPI分别存放在三张表中](#二--宏ReadGIBIPI导入excel并将GIBIPI分别存放在三张表中)   
* [三. 宏ProGIBIPI对单个GIBIPI数据集进行处理](#三--宏ProGIBIPI对单个GIBIPI数据集进行处理)    
* [四. 将GIBIPI结果综合至一张表](#四--将GIBIPI结果综合至一张表)  
 
    
&ensp;&ensp;&ensp;&ensp;  


# 一  原始数据  
&ensp;&ensp;&ensp;&ensp;    
&ensp;&ensp;&ensp;&ensp;牙龈炎临床试验中GIBIPI指标原始excel数据如下：Origin.xlsx    
&ensp;&ensp;&ensp;&ensp;  
![image](https://github.com/TracyHuo/GIBIPIProcess/blob/master/Image/1.PNG);  
&ensp;&ensp;&ensp;&ensp;   
&ensp;&ensp;&ensp;&ensp;   
# 二  宏ReadGIBIPI导入excel并将GIBIPI分别存放在三张表中   
&ensp;&ensp;&ensp;&ensp;    
* **代码**：  
&ensp;&ensp;&ensp;&ensp;    
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
&ensp;&ensp;&ensp;&ensp;   
&ensp;&ensp;&ensp;&ensp;    
* **结果**：    
&ensp;&ensp;&ensp;&ensp;   
&ensp;&ensp;&ensp;&ensp;原始数据集origin  
&ensp;&ensp;&ensp;&ensp;   
![image](https://github.com/TracyHuo/GIBIPIProcess/blob/master/Image/origin.PNG);  
&ensp;&ensp;&ensp;&ensp;   
&ensp;&ensp;&ensp;&ensp;   
&ensp;&ensp;&ensp;&ensp;以GI指标为例：  
&ensp;&ensp;&ensp;&ensp;GI  
&ensp;&ensp;&ensp;&ensp;  
![image](https://github.com/TracyHuo/GIBIPIProcess/blob/master/Image/GI.PNG);  
&ensp;&ensp;&ensp;&ensp;   
&ensp;&ensp;&ensp;&ensp;      
&ensp;&ensp;&ensp;&ensp;    
&ensp;&ensp;&ensp;&ensp;   
# 三  宏ProGIBIPI对单个GIBIPI数据集进行处理   
&ensp;&ensp;&ensp;&ensp;   
&ensp;&ensp;&ensp;&ensp;    
* **代码**：   
&ensp;&ensp;&ensp;&ensp;    
```
/*****************************************************************************************
此宏的作用是对单个的GI/BI/PI数据集进行处理，以GI为例，处理mylib.GI，返回数据集mylib.GI_pro_5
其列为：ID,Name,Toothposition，GI. 其中，Toothposition是牙位，GI是数值型变量。
BI,PI同理。
*******************************************************************************************/

%MACRO ProGIBIPI(dataset= );
    %let Index=%scan(&dataset, 2);

	/********************************************************************
     1.在mylib.GI表中，一位受试者的数据GI数据占8行，每行21个值，共168个值。
	 以下代码的目的是把这8行数据整合成一行，并赋予col_1-col_168的列名。
	 得到的 mylib.GI_pro_1的列名为：Name ID col_1-col_168 .每位受试者对应
	 一行数据。
	 BI,PI同理。此宏的dataset参数可以决定处理mylib.GI，mylib.BI，mylib.PI
	 中的哪张表。
	 以下代码，创建了col_1-col_168的变量，并用数组管理。linenum=mod(_N_,8)
	 实现了类似循环的方式，以8行为单位处理（每位受试者的GI数据占8行），
     newcol{(linenum-1)*21+i}=oldcol{i}是把原本的一行GI数据赋给col_1-col_168
	 里对应的新变量。
	********************************************************************/

    DATA &dataset._pro_1(keep=Name ID col_1-col_168); /*一定注意&dataset后的点号，一定要有！
	                                                    否则报错*/
        Length Name $ 9 ID $ 2 ;/*这句主要是为了调整PDV里变量顺序*/
        Retain col_1-col_168  ; /*一定要retain*/
        ARRAY newcol{168} $ col_1-col_168;
        SET &dataset;
        ARRAY oldcol{21} $ C_1--C_21; /*要写在SET后边，否则提示找不到变量C_1*/

        linenum=mod(_N_,8);
        IF linenum NE 0 THEN DO;
                                DO i=1 TO 21;
                                    newcol{(linenum-1)*21+i}=oldcol{i};
                                END;
                             END;

        IF linenum=0 THEN DO;
                             linenum=8 ;
                             DO i=1 TO 21;
                                 newcol{(linenum-1)*21+i}=oldcol{i};
                             END;
                             OUTPUT;
							 /*只在此处用OUTPUT，这个很重要，保证了只输出
							 col_1到col_168都赋了值的那一行。注意col_1到
							 col_168通过retain语句retain*/
                          END;

						/*本来想写IF linemun=0 THEN column{1}-column{21} = C_1--C_21 ;
                        IF linemun=2 THEN col_22-col_42 = C_1--C_21 ;
                        IF linemun=3 THEN col_43-col_63 = C_1--C_21 ;
                        IF linemun=4 THEN col_64-col_84 = C_1--C_21 ;
                        IF linemun=5 THEN col_85-col_105 = C_1--C_21 ;
                        IF linemun=6 THEN col_106-col_126 = C_1--C_21 ;
                        IF linemun=7 THEN col_127-col_147 = C_1--C_21 ;
                        IF linemun=0 THEN DO;col_148-col_168 = C_1--C_21 ; OUTPUT; END; 
						但是报错，应该是THEN后边不能用变量列表。*/

    RUN;

    /**************************************************************************************
	 2.转置。mylib.GI_pro_1里的col_1-col_168是列，每位受试者一行数据。以下代码将col_1-col_168列
	 转置，这样，每位受试者的GI数据对应1列，168行。得到的mylib.GI_pro_3里，共ID,Name,_NAME_,GI
	 四列。其中，_NAME_是系统自动生成的。
	***************************************************************************************/
    PROC SORT data=&dataset._pro_1  out=&dataset._pro_2;
    BY ID Name;
    RUN;

    PROC TRANSPOSE data=&dataset._pro_2  out=&dataset._pro_3(rename=(COL1 = &Index));/*COL_1是系统自取的名字*/
    BY ID Name;
    VAR col_1-col_168 ;
    RUN;

    /*****************************************************************************
	 3. 改_name_列值。mylib.GI_pro_3里的_NAME_是系统自动生成的列，里面的值是col_1-col_168。
	 而事实上，应该指明每个GI值所对应的牙位。所以，需要新建列Toothposition，指明每个GI值对应
	 的牙位。可以用以下代码实现，得到mylib.GI_pro_4，包含ID,Name,GI,Toothposition四列。
	 其中，Toothposition是类似 17_DB的牙位。
	*****************************************************************************/

    DATA &dataset._pro_4(keep=Name ID Toothposition &Index);
	/*注意，数据集选项执行顺序：drop>keep>retain*/

        /*以下定义的数组和三重循环是为了实现这样的目的：创建一个含168个字符元素tp_1 - tp_168的数组
	     toothp，然后，把牙位名如 17_DB 赋给其中的元素，得到168个牙位名，存储到tp_1 - tp_168变量里。*/
        Retain tp_1 - tp_168 ;
        IF _N_=1 THEN 
		/*这个很重要，也就是只在第一轮DATA循环的时候赋值toothposition数组。以免每次DATA步都赋一遍数组*/
        DO;         
            ARRAY toothB{3} $ ("DB","B","MB");
            ARRAY toothL{3} $ ("DL","L","ML");
            ARRAY toothp{168}$ 8 tp_1 - tp_168 ; 
			/*虽然只在_N_=1时创建数组，但是数组是整个DATA步里都能使用的。又因为我retain了数组变量，
			所以每一轮DATA步都可以用数组，都有值*/

            DO i=1 TO 8 ;
                DO j=7 TO 1 by -1; /*一定注意步长是-1，若不写BY则不会报错，但也不会有结果*/
                    DO k=1 TO 3 ;
                        var1 = ceil(i/2)*10 + j;
                        count= (i-1)*21 + (7-j)*3 +k; /*count是现在循环进行的轮数*/
                        IF      mod(i,2) NE 0 THEN toothp{count}=catx("_", var1, toothB{k}) ; 
						/*i是奇数*/
                        ELSE IF mod(i,2) =0 THEN toothp{count}=catx("_", var1, toothL{k}) ;
						/*i是偶数*/
                    END;
                END;
            END;
         END;


        SET &dataset._pro_3;
        /*linenum=mod(_N_, 168)实现类似循环的效果。因为一位受试者的数据占168行，所以以168行为单位
		 进行处理。因为已经创建了tp_1 - tp_168这168个变量，含有各个牙位的名字，用toothp数组管理，
		 所以可以很方便得赋值给Toothposition牙位变量。*/
        linenum=mod(_N_, 168);
        IF linenum NE 0 THEN Toothposition = toothp{linenum};
        ELSE IF linenum=0 THEN Toothposition = toothp{168};

    RUN;
    
	/*****************************************************************************************
     4.把GI列改为数值型。因为原始excel的列是混合类型，所以PROC IMPORT导入SAS后都是字符型列。而GI值
	 本身是数值型的，所以以下代码使用input函数进行了格式转换。得到的得到mylib.GI_pro_5，
	 包含ID,Name,Toothposition，GI四列，GI是数值型的。而PI/BI同理。
	*****************************************************************************************/
	DATA &dataset._pro_5;
	    SET &dataset._pro_4;
        Index2 = input(&Index,1.);
        DROP &Index ;
		RENAME Index2=&Index;
	RUN;


%MEND ProGIBIPI ;

%ProGIBIPI(dataset=mylib.GI)
%ProGIBIPI(dataset=mylib.BI)
%ProGIBIPI(dataset=mylib.PI)

```   
* **结果**：    
&ensp;&ensp;&ensp;&ensp;   
&ensp;&ensp;&ensp;&ensp;以GI指标为例：  
&ensp;&ensp;&ensp;&ensp;GI_pro_1  
&ensp;&ensp;&ensp;&ensp;  
![image](https://github.com/TracyHuo/GIBIPIProcess/blob/master/Image/GI_pro_1.PNG);  
&ensp;&ensp;&ensp;&ensp;   
&ensp;&ensp;&ensp;&ensp;      
&ensp;&ensp;&ensp;&ensp;GI_pro_3  
&ensp;&ensp;&ensp;&ensp;   
![image](https://github.com/TracyHuo/GIBIPIProcess/blob/master/Image/GI_pro_3.PNG);  
&ensp;&ensp;&ensp;&ensp;   
&ensp;&ensp;&ensp;&ensp;    
&ensp;&ensp;&ensp;&ensp;GI_pro_5  
&ensp;&ensp;&ensp;&ensp;   
![image](https://github.com/TracyHuo/GIBIPIProcess/blob/master/Image/GI_pro_5.PNG);  
&ensp;&ensp;&ensp;&ensp;   
&ensp;&ensp;&ensp;&ensp;    
&ensp;&ensp;&ensp;&ensp;   
# 四  将GIBIPI结果综合至一张表    
&ensp;&ensp;&ensp;&ensp;   
&ensp;&ensp;&ensp;&ensp;    
* **代码**：   
&ensp;&ensp;&ensp;&ensp;    
```
DATA mylib.result;
MERGE mylib.GI_pro_5 mylib.BI_pro_5 mylib.PI_pro_5 ;
RUN;
```
* **结果**：    
&ensp;&ensp;&ensp;&ensp;     
&ensp;&ensp;&ensp;&ensp;result  
&ensp;&ensp;&ensp;&ensp;  
![image](https://github.com/TracyHuo/GIBIPIProcess/blob/master/Image/result.PNG);  
&ensp;&ensp;&ensp;&ensp;   
&ensp;&ensp;&ensp;&ensp;      
