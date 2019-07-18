//
//  ViewController.m
//  ExcelParse
//
//  Created by 朱志佳 on 2019/7/18.
//  Copyright © 2019 朱志佳. All rights reserved.
//

#import "ViewController.h"

#import "LAWExcelTool.h"

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
}

@end
