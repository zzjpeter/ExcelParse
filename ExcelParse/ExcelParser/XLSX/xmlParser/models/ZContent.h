/******************************************************/
//
//  ProjectName: ZXLSXReader
//  FileName   : ZContent.h
//  Author     : Casanova.Z/朱静宁 16/8/25.
//  E-mail     : casanova.z@qq.com
//  Blog       : http://blog.sina.com.cn/casanovaZHU
//
//  Copyright © 2016年 casanova. All rights reserved.
/******************************************************/

#import <Foundation/Foundation.h>

@interface ZContent : NSObject<NSCoding>

@property (nonatomic, copy) NSString *sheetName;    // 所在文件的名称(sheet1,sheet2)
@property (nonatomic, copy) NSString *keyName;      // 在sheet中的标示(A1,B3)
@property (nonatomic, copy) NSString *value;        // 对应文件下对应标识下的值

//从keyName中获取[通过首字母区分列数，后面的数字区分行数]
/*
 keyName数据解析：
 行数解析: A1:代表第1行 A2:代表第2行 A3:代表第3行 依次类推。。。
 列数解析: A代表第1列 B:代表第2列 C:代表第3列
 */
@property (nonatomic,copy) NSString *column;//列数
//@property (nonatomic,assign) NSInteger columnNum;//列数
//
//@property (nonatomic,copy) NSString *row;//行数
//@property (nonatomic,assign) NSInteger rowNum;//行数

@end
