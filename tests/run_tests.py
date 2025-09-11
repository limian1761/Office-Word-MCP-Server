# -*- coding: utf-8 -*-
"""
测试运行器脚本
用于统一执行所有测试并提供友好的测试结果报告
"""

import os
import sys
import unittest
import argparse
import time
from datetime import datetime
import importlib.util
from types import ModuleType
from typing import List, Tuple, Dict, Optional


class ColoredTestResult(unittest.TextTestResult):
    """自定义测试结果输出，添加颜色和更详细的信息"""

    # ANSI颜色代码
    COLORS = {
        'HEADER': '\033[95m',
        'OKBLUE': '\033[94m',
        'OKGREEN': '\033[92m',
        'WARNING': '\033[93m',
        'FAIL': '\033[91m',
        'ENDC': '\033[0m',
        'BOLD': '\033[1m',
        'UNDERLINE': '\033[4m'
    }

    def __init__(self, stream, descriptions, verbosity):
        super().__init__(stream, descriptions, verbosity)
        self.success_count = 0
        self.test_start_times = {}
        self.test_durations = {}

    def startTest(self, test):
        """记录测试开始时间"""
        self.test_start_times[test.id()] = time.time()
        super().startTest(test)

    def addSuccess(self, test):
        """记录成功的测试"""
        duration = time.time() - self.test_start_times[test.id()]
        self.test_durations[test.id()] = duration
        self.success_count += 1
        if self.showAll:
            self.stream.write(self.COLORS['OKGREEN'])
            self.stream.write("✓ " + self.getDescription(test) + " ")
            self.stream.write(self.COLORS['ENDC'])
            self.stream.write(f"({duration:.3f}s)\n")
        elif self.dots:
            self.stream.write(self.COLORS['OKGREEN'])
            self.stream.write('.')
            self.stream.write(self.COLORS['ENDC'])
            self.stream.flush()

    def addFailure(self, test, err):
        """记录失败的测试"""
        duration = time.time() - self.test_start_times[test.id()]
        self.test_durations[test.id()] = duration
        if self.showAll:
            self.stream.write(self.COLORS['FAIL'])
            self.stream.write("✗ " + self.getDescription(test) + " ")
            self.stream.write(self.COLORS['ENDC'])
            self.stream.write(f"({duration:.3f}s)\n")
        elif self.dots:
            self.stream.write(self.COLORS['FAIL'])
            self.stream.write('F')
            self.stream.write(self.COLORS['ENDC'])
            self.stream.flush()
        self.failures.append((test, err))

    def addError(self, test, err):
        """记录出错的测试"""
        duration = time.time() - self.test_start_times[test.id()]
        self.test_durations[test.id()] = duration
        if self.showAll:
            self.stream.write(self.COLORS['WARNING'])
            self.stream.write("! " + self.getDescription(test) + " ")
            self.stream.write(self.COLORS['ENDC'])
            self.stream.write(f"({duration:.3f}s)\n")
        elif self.dots:
            self.stream.write(self.COLORS['WARNING'])
            self.stream.write('E')
            self.stream.write(self.COLORS['ENDC'])
            self.stream.flush()
        # 确保错误信息是正确的格式
        self.errors.append((test, self._format_error_message(err)))

    def printSummary(self):
        """打印测试摘要"""
        self.stream.writeln()
        self.stream.writeln("=" * 80)
        
        # 计算总测试数和总耗时
        total_tests = self.success_count + len(self.failures) + len(self.errors)
        total_duration = sum(self.test_durations.values())
        
        # 打印测试统计信息
        self.stream.write(f"{self.COLORS['HEADER']}测试统计信息:{self.COLORS['ENDC']}\n")
        self.stream.write(f"总测试数: {total_tests}\n")
        self.stream.write(f"通过: {self.COLORS['OKGREEN']}{self.success_count}{self.COLORS['ENDC']}\n")
        self.stream.write(f"失败: {self.COLORS['FAIL']}{len(self.failures)}{self.COLORS['ENDC']}\n")
        self.stream.write(f"错误: {self.COLORS['WARNING']}{len(self.errors)}{self.COLORS['ENDC']}\n")
        self.stream.write(f"总耗时: {total_duration:.3f}s\n")
        
        # 打印失败的测试详情
        if self.failures:
            self.stream.write(f"\n{self.COLORS['FAIL']}失败的测试详情:{self.COLORS['ENDC']}\n")
            for i, (test, err) in enumerate(self.failures, 1):
                self.stream.write(f"{i}. {test.id()}:\n")
                self.stream.write(f"   {str(err[1])}\n")
        
        # 打印出错的测试详情
        if self.errors:
            self.stream.write(f"\n{self.COLORS['WARNING']}出错的测试详情:{self.COLORS['ENDC']}\n")
            for i, (test, err) in enumerate(self.errors, 1):
                self.stream.write(f"{i}. {test.id()}:\n")
                self.stream.write(f"   {str(err[1])}\n")
        
        self.stream.writeln("=" * 80)
        
        # 根据测试结果返回状态码
        if not self.wasSuccessful():
            self.stream.write(f"{self.COLORS['FAIL']}测试未通过!{self.COLORS['ENDC']}\n")
        else:
            self.stream.write(f"{self.COLORS['OKGREEN']}所有测试通过!{self.COLORS['ENDC']}\n")

    def _format_error_message(self, err):
        """格式化错误信息，确保它是一个字符串"""
        if isinstance(err, tuple):
            # 标准的unittest错误元组 (type, value, traceback)
            return str(err[1])
        elif hasattr(err, '__str__'):
            return str(err)
        else:
            return repr(err)

    def wasSuccessful(self):
        """检查测试是否全部通过"""
        return len(self.failures) == 0 and len(self.errors) == 0
        
    def printErrorList(self, flavour, errors):
        """重写printErrorList方法，使用更安全的格式化方式处理错误信息"""
        for test, err in errors:
            self.stream.write(self.COLORS['FAIL'])
            self.stream.writeln(f"{flavour}: {self.getDescription(test)}")
            self.stream.write(self.COLORS['ENDC'])
            # 使用更安全的方式格式化错误信息
            if isinstance(err, tuple) and len(err) >= 2:
                err_msg = str(err[1])
            elif hasattr(err, '__str__'):
                err_msg = str(err)
            else:
                err_msg = repr(err)
            self.stream.writeln(err_msg)
            self.stream.writeln()


class ColoredTestRunner(unittest.TextTestRunner):
    """自定义测试运行器，使用带颜色的测试结果输出"""

    def _makeResult(self):
        """创建测试结果对象"""
        return ColoredTestResult(self.stream, self.descriptions, self.verbosity)

    def run(self, test):
        """运行测试并输出结果"""
        result = super().run(test)
        # 打印测试摘要
        if hasattr(result, 'printSummary'):
            result.printSummary()
        return result


def find_test_files(test_dir: str) -> List[str]:
    """查找指定目录下所有的测试文件"""
    test_files = []
    # 遍历目录下所有文件
    for file_name in os.listdir(test_dir):
        # 检查是否是测试文件（以test_开头，以.py结尾）
        if file_name.startswith('test_') and file_name.endswith('.py'):
            # 排除主测试运行器脚本
            if file_name != 'run_tests.py':
                test_files.append(os.path.join(test_dir, file_name))
    return test_files


def load_test_module(file_path: str) -> Optional[ModuleType]:
    """加载测试模块"""
    try:
        # 获取模块名称
        module_name = os.path.splitext(os.path.basename(file_path))[0]
        # 创建模块规范
        spec = importlib.util.spec_from_file_location(module_name, file_path)
        if spec is None:
            print(f"无法创建模块规范: {file_path}")
            return None
        # 创建模块
        module = importlib.util.module_from_spec(spec)
        # 执行模块代码
        spec.loader.exec_module(module)
        return module
    except Exception as e:
        print(f"加载测试模块失败 {file_path}: {str(e)}")
        return None


def collect_tests(test_dir: str, test_pattern: Optional[str] = None, exclude_pattern: Optional[str] = None) -> unittest.TestSuite:
    """收集测试用例"""
    # 创建测试套件
    test_suite = unittest.TestSuite()
    
    # 查找测试文件
    test_files = find_test_files(test_dir)
    
    # 按名称排序，确保测试执行顺序一致
    test_files.sort()
    
    # 遍历测试文件
    for file_path in test_files:
        file_name = os.path.basename(file_path)
        
        # 检查是否需要排除
        if exclude_pattern and exclude_pattern in file_name:
            print(f"排除测试文件: {file_name}")
            continue
        
        # 检查是否匹配测试模式
        if test_pattern and test_pattern not in file_name:
            continue
        
        # 加载测试模块
        module = load_test_module(file_path)
        if module is None:
            continue
        
        # 查找模块中的测试用例
        tests = unittest.defaultTestLoader.loadTestsFromModule(module)
        if tests.countTestCases() > 0:
            test_suite.addTests(tests)
            print(f"加载测试文件: {file_name} ({tests.countTestCases()}个测试)")
        else:
            print(f"跳过空测试文件: {file_name}")
    
    return test_suite


def run_tests(test_dir: str, verbose: bool = False, test_pattern: Optional[str] = None, exclude_pattern: Optional[str] = None) -> bool:
    """运行测试并返回是否全部通过"""
    # 收集测试用例
    test_suite = collect_tests(test_dir, test_pattern, exclude_pattern)
    
    # 检查是否有测试用例
    if test_suite.countTestCases() == 0:
        print(f"{ColoredTestResult.COLORS['WARNING']}未找到测试用例!{ColoredTestResult.COLORS['ENDC']}")
        return False
    
    print(f"\n{ColoredTestResult.COLORS['HEADER']}开始执行测试...{ColoredTestResult.COLORS['ENDC']}")
    print(f"总共找到 {test_suite.countTestCases()} 个测试用例")
    print(f"测试开始时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 80)
    
    # 创建测试运行器
    runner = ColoredTestRunner(verbosity=2 if verbose else 1, stream=sys.stdout)
    
    # 运行测试
    result = runner.run(test_suite)
    
    # 返回测试是否全部通过
    return result.wasSuccessful()


def main():
    """主函数"""
    # 创建参数解析器
    parser = argparse.ArgumentParser(description='运行Word文档处理工具的测试用例')
    parser.add_argument('-v', '--verbose', action='store_true', help='显示详细的测试输出')
    parser.add_argument('-t', '--test-pattern', help='只运行名称包含指定模式的测试文件')
    parser.add_argument('-e', '--exclude', help='排除名称包含指定模式的测试文件')
    parser.add_argument('-d', '--test-dir', default=os.path.dirname(os.path.abspath(__file__)), help='测试文件所在目录')
    
    # 解析命令行参数
    args = parser.parse_args()
    
    # 确保测试目录存在
    if not os.path.isdir(args.test_dir):
        print(f"错误: 测试目录不存在: {args.test_dir}")
        sys.exit(1)
    
    # 添加项目根目录到Python路径，确保可以导入模块
    project_root = os.path.dirname(args.test_dir)
    if project_root not in sys.path:
        sys.path.insert(0, project_root)
    
    # 运行测试
    success = run_tests(args.test_dir, args.verbose, args.test_pattern, args.exclude)
    
    # 根据测试结果设置退出码
    sys.exit(0 if success else 1)


if __name__ == '__main__':
    main()