//
//  ViewController.m
//  ExcelParse
//
//  Created by 朱志佳 on 2019/7/18.
//  Copyright © 2019 朱志佳. All rights reserved.
//

#import "ViewController.h"

#import "LAWExcelTool.h"
#import "ZContent.h"

@interface ViewController ()<LAWExcelParserDelegate>

@end

@implementation ViewController

- (void)viewDidLoad {
    [super viewDidLoad];
    // Do any additional setup after loading the view.
    [self testParse];
}

- (void)testParse
{
    [LAWExcelTool shareInstance].delegate = self;
}

- (IBAction)csvParse:(id)sender {
    NSString *path = [[NSBundle mainBundle] pathForResource:@"test.csv" ofType:nil];
    [[LAWExcelTool shareInstance] parserExcelWithPath:path];
}

- (IBAction)xlsParse:(id)sender {
    NSString *path = [[NSBundle mainBundle] pathForResource:@"test.xls" ofType:nil];
    [[LAWExcelTool shareInstance] parserExcelWithPath:path];
}

- (IBAction)xlsxParse:(id)sender {
    NSString *path = [[NSBundle mainBundle] pathForResource:@"test3.xlsx" ofType:nil];
    [[LAWExcelTool shareInstance] parserExcelWithPath:path];
}

#pragma mark LAWExcelParserDelegate
- (void)parser:(LAWExcelTool *)parser success:(id)responseObj
{
    NSLog(@"responseObj:%@", responseObj);
    [self handle:responseObj];
}

#pragma mark handle
- (void)handle:(NSMutableArray<ZContent *> *)responseObj
{
    if (![responseObj isKindOfClass:[NSArray class]] || !(responseObj.count > 0)) {
        return;
    }
    
    NSMutableArray<NSArray *> *bigArr = [self handleToRowWithArray:responseObj];
    
    [self handleRowArrayToContentKeyValueStr:bigArr];

}

#pragma mark 处理成行
- (NSMutableArray *)handleToRowWithArray:(NSMutableArray<ZContent *> *)responseObj
{
    ZContent *lastContent = responseObj.lastObject;
    NSString *lastColumn = lastContent.column;
    
    NSMutableArray *bigArr = [NSMutableArray new];
    NSMutableArray *rowArr = [NSMutableArray new];
    for (ZContent *content in responseObj) {
        if ([content.column isEqualToString:lastColumn]) {//一行结束
            [rowArr addObject:content];
            [bigArr addObject:rowArr];
            rowArr = [NSMutableArray new];
            continue;
        }
        [rowArr addObject:content];
    }
    NSLog(@"%@",bigArr);
    
    return bigArr;
}

#pragma mark 处理成键值对字符串
- (void )handleRowArrayToContentKeyValueStr:(NSMutableArray<NSArray *> *)bigArr
{
    NSMutableArray<NSMutableString *> *bigContentStrArr = [NSMutableArray new];
    NSInteger column = bigArr.firstObject.count;
    for (NSInteger i = 0; i < column; i++) {
        NSMutableString *contentStr = [NSMutableString new];
        [bigContentStrArr addObject:contentStr];
    }
    
    for (NSArray<ZContent *> *rowArr in bigArr) {//
        if (rowArr.count <= 1) {
            return ;
        }
        NSString *key = rowArr.firstObject.value;
        for (NSInteger index = 1; index < rowArr.count; index++) {
            NSString *value = rowArr[index].value;
            NSMutableString *content = [NSMutableString stringWithFormat:@"%@ =  %@;\n ",key, value];
            [bigContentStrArr[index] appendString:content];
        }
    }
    
    NSLog(@"bigContentStrArr:%@",bigContentStrArr);
}

@end
