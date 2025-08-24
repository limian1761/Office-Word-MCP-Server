#!/usr/bin/env python3
"""
兼容入口文件 - 已统一指向官方启动入口

注意：此文件仅作为向后兼容使用，实际调用的是word_document_server/main.py中的run_server()函数
项目的官方启动入口为word_document_server/main.py
"""
from word_document_server.main import run_server

if __name__ == "__main__":
    run_server()