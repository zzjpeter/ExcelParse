//
//  QuickExcelReaderUtil.m
//  QuickExcelKit
//
//  Created by pcjbird on 2018/3/17.
//  Copyright © 2018年 Zero Status. All rights reserved.
//

#import "QuickExcelReaderUtil.h"
#import "ZWorkbookParser.h"
#import "ZXLSXParser.h"
#import "ZContent.h"
#import "CSV/CSVParser.h"
#import "DHxlsReaderIOS.h"
#import "ZHeader.h"

#define QUICKEXCELKIT_ERROR(ecode, msg)  [NSError errorWithDomain:@"QuickExcelReaderUtil" code:(ecode) userInfo:([NSDictionary dictionaryWithObjectsAndKeys:(msg), @"message", nil])]

@interface QuickExcelReaderUtil()<ZXLSXParserDelegate>

@property (nonatomic, strong) ZXLSXParser *xmlPaser;

@property (nonatomic, strong) NSMutableDictionary<NSString*, QuickExcelReaderBlock>*callbacks;
@end

@implementation QuickExcelReaderUtil

static QuickExcelReaderUtil *instance;
+ (instancetype)sharedManager{
    static dispatch_once_t onceToken;
    dispatch_once(&onceToken, ^{
        instance = [self new];
    });
    return instance;
}

//初始化 默认配置
- (instancetype)init{
    if(self = [super init]){
        self.xmlPaser = [ZXLSXParser defaultZXLSXParser];
        [self.xmlPaser setParseOutType:ZParseOutTypeArrayObj];
        self.xmlPaser.delegate = self;
        self.callbacks = [NSMutableDictionary<NSString*, QuickExcelReaderBlock> dictionary];
    }
    return self;
}

// api
+ (void)readExcelWithPath:(NSString*)filePath complete:(QuickExcelReaderBlock)block
{
    NSString *mimeType = [[[filePath componentsSeparatedByString:@"."] lastObject] lowercaseString];
    
    if([mimeType isEqualToString:@"csv"]){
        [QuickExcelReaderUtil.sharedManager parserExcel_CSV_WithPath:filePath complete:block];
    }
    else if ([mimeType isEqualToString:@"xls"]){
        [QuickExcelReaderUtil.sharedManager parserExcel_XLS_WithPath:filePath complete:block];
    }
    else if ([mimeType isEqualToString:@"xlsx"]){
        [QuickExcelReaderUtil.sharedManager parserExcel_XLSX_WithPath:filePath complete:block];
    }
    else{
        NSString *errorString = [NSString stringWithFormat:@"读取 Excel 失败，格式 %@ 不支持。", mimeType];
        !block ? : block(nil, QUICKEXCELKIT_ERROR(1000, errorString));
    }
}

#pragma mark 解析XLSX类型的excel
-(void)parserExcel_XLSX_WithPath:(NSString *)filePath complete:(QuickExcelReaderBlock)block{
    if(block){
        [self.callbacks removeAllObjects];
        [self.callbacks setObject:block forKey:filePath];
        [self.xmlPaser setParseFilePath:filePath];
        self.xmlPaser.delegate = self;
        [self.xmlPaser parse];
    }
}
#pragma mark ZXLSXParserDelegate
- (void)parser:(ZXLSXParser *)parser success:(id)responseObj
{
    NSString *filePath = parser.parseFilePath;
    QuickExcelReaderBlock block = [self.callbacks objectForKey:filePath];
    if(![responseObj isKindOfClass:[NSArray<ZContent *> class]])
    {
        NSString *errorString = [NSString stringWithFormat:@"读取 Excel: %@ 失败。", [filePath lastPathComponent]];
        if(block)block(nil, QUICKEXCELKIT_ERROR(1000, errorString));
        [self.callbacks removeAllObjects];
        return;
    }
    NSMutableDictionary<NSString*, NSArray<ZContent*>*>* results = [NSMutableDictionary<NSString*, NSArray<ZContent*>*> dictionary];
    NSString *defaultSheetName = @"xlsx";
    NSString *defaultKeyName = @"xlsx_key";
    for (ZContent * content in responseObj) {
        if(!content.sheetName) content.sheetName = defaultSheetName;
        if(!content.keyName) content.keyName = defaultKeyName;
        NSMutableArray<ZContent *> *resultArray = (NSMutableArray<ZContent *> *)[results objectForKey:content.sheetName];
        if(![resultArray isKindOfClass:[NSMutableArray<ZContent *> class]])
        {
            resultArray = [NSMutableArray array];
            [results setObject:resultArray forKey:content.sheetName];
        }
        [resultArray addObject:content];
    }
    if(block)block(results, nil);
    [self.callbacks removeAllObjects];
}

#pragma mark 解析XLS类型的excel
-(void)parserExcel_XLS_WithPath:(NSString *)filePath complete:(QuickExcelReaderBlock)block
{
    DHxlsReader *reader = [DHxlsReader xlsReaderFromFile:filePath];
    if(![reader isKindOfClass:[DHxlsReader class]]){
        NSString *errorString = [NSString stringWithFormat:@"读取 Excel: %@ 失败。", [filePath lastPathComponent]];
        if(block)block(nil, QUICKEXCELKIT_ERROR(1000, errorString));
        return;
    }
   
    NSMutableDictionary<NSString*, NSArray<ZContent*>*>* results = [NSMutableDictionary<NSString*, NSArray<ZContent*>*> dictionary];
    NSInteger sheetsCount = [reader numberOfSheets];
    for (uint32_t i = 0; i < sheetsCount; i++){
        NSString *sheetName = [reader sheetNameAtIndex:i];
        NSMutableArray<ZContent *> *resultArray = [NSMutableArray array];
        [reader startIterator:i];
        int rows = reader.getRows;
        int cols = reader.getCols;
        for(int r = 1; r <= rows; r++){
            for(int c = 1;c <= cols; c++){
                unichar ch =64 + c;
                NSString *str =[NSString stringWithUTF8String:(char *)&ch];
                DHcell *cell = [reader cellInWorkSheetIndex:i row:r col:c];
                ZContent *content = [[ZContent alloc] init];
                content.sheetName = [reader sheetNameAtIndex:0];
                content.keyName = [NSString stringWithFormat: @"%@%d", str,c+r-1];
                content.value = [cell dump];
                [resultArray addObject:content];
            }
        }
        [results setObject:resultArray forKey:sheetName];
    }
    if(block)block(results, nil);
}
#pragma mark 解析CSV类型的excel
-(void)parserExcel_CSV_WithPath:(NSString *)filePath complete:(QuickExcelReaderBlock)block
{
    NSMutableArray *array = [CSVParser readCSVData:filePath];
    if(![array isKindOfClass:[NSArray class]]){
        NSString *errorString = [NSString stringWithFormat:@"读取 Excel: %@ 失败。", [filePath lastPathComponent]];
        if(block)block(nil, QUICKEXCELKIT_ERROR(1000, errorString));
        return;
    }
    NSMutableDictionary<NSString*, NSArray<ZContent*>*>* results = [NSMutableDictionary<NSString*, NSArray<ZContent*>*> dictionary];
    NSString *sheetName = @"csv";
    NSMutableArray<ZContent *> *resultArray = [NSMutableArray array];
    int row = 0;
    for(NSArray *item in array){
        for(int i = 0; i < item.count; i++) {
            unichar ch =65 + i;
            NSString *str =[NSString stringWithUTF8String:(char *)&ch];
            ZContent *content = [[ZContent alloc] init];
            content.sheetName = sheetName;
            content.keyName = [NSString stringWithFormat: @"%@%d", str,i+1+row];
            content.value = item[i];
            [resultArray addObject:content];
        }
        row++;
    }
    [results setObject:resultArray forKey:sheetName];
    if(block)block(results, nil);
}

@end
