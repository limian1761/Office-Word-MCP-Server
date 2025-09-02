# 工具持久化执行指南

确保A/B工具不间断执行且不会因调用过多而被终止的解决方案。

## 1. 心跳机制

实现心跳机制以保持工具活跃状态：

```python
# 心跳检测示例
import time
import threading

def heartbeat():
    while True:
        # 发送心跳信号
        print("Heartbeat signal sent")
        time.sleep(30)  # 每30秒发送一次心跳

# 在工具启动时运行心跳线程
heartbeat_thread = threading.Thread(target=heartbeat, daemon=True)
heartbeat_thread.start()
```

## 2. 错误恢复机制

实现错误恢复和重试机制：

```python
import time
from functools import wraps

def retry_on_failure(max_retries=3, delay=1):
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            for attempt in range(max_retries):
                try:
                    return func(*args, **kwargs)
                except Exception as e:
                    if attempt == max_retries - 1:
                        raise e
                    time.sleep(delay * (2 ** attempt))  # 指数退避
            return None
        return wrapper
    return decorator

@retry_on_failure(max_retries=5, delay=2)
def critical_operation():
    # 可能失败的关键操作
    pass
```

## 3. 资源管理

合理管理系统资源防止内存泄漏：

```python
import gc
import psutil
import os

def monitor_resources():
    # 检查内存使用情况
    process = psutil.Process(os.getpid())
    memory_usage = process.memory_info().rss / 1024 / 1024  # MB
    
    if memory_usage > 500:  # 如果内存使用超过500MB
        gc.collect()  # 强制垃圾回收
        
    return memory_usage

# 在长时间运行的循环中定期调用
def long_running_task():
    for i in range(1000000):
        # 执行任务
        if i % 1000 == 0:
            memory_usage = monitor_resources()
            print(f"Memory usage: {memory_usage} MB")
```

## 4. 连接保持

对于需要维持连接的工具，实现连接保持机制：

```python
import time
import socket

class PersistentConnection:
    def __init__(self, host, port):
        self.host = host
        self.port = port
        self.socket = None
        self.connect()
    
    def connect(self):
        try:
            self.socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.socket.connect((self.host, self.port))
        except Exception as e:
            print(f"Connection failed: {e}")
    
    def keep_alive(self):
        while True:
            try:
                # 发送保持连接的信号
                self.socket.send(b'PING')
                response = self.socket.recv(1024)
                if not response:
                    raise ConnectionError("Connection lost")
            except Exception as e:
                print(f"Connection error: {e}")
                self.connect()  # 重新连接
            time.sleep(60)  # 每分钟发送一次心跳
```

## 5. 日志和监控

实现全面的日志记录和监控：

```python
import logging
from datetime import datetime

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('tool_execution.log'),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger('PersistentTool')

def log_execution_status(status, details=""):
    logger.info(f"Execution status: {status}. Details: {details}")

# 在工具关键节点添加日志
def tool_main_function():
    logger.info("Tool started")
    try:
        # 主要功能逻辑
        pass
        logger.info("Tool completed successfully")
    except Exception as e:
        logger.error(f"Tool failed with error: {e}")
        raise
```

## 6. 分布式执行

对于长时间运行的任务，考虑使用分布式执行：

```python
import multiprocessing
import queue
import time

def worker(task_queue, result_queue):
    while True:
        try:
            task = task_queue.get(timeout=1)
            if task is None:
                break
            
            # 执行任务
            result = process_task(task)
            result_queue.put(result)
        except queue.Empty:
            continue
        except Exception as e:
            result_queue.put(('error', str(e)))

def process_task(task):
    # 实际的任务处理逻辑
    return f"Processed {task}"

# 使用示例
if __name__ == "__main__":
    task_queue = multiprocessing.Queue()
    result_queue = multiprocessing.Queue()
    
    # 启动工作进程
    processes = []
    for i in range(4):  # 4个工作进程
        p = multiprocessing.Process(target=worker, args=(task_queue, result_queue))
        p.start()
        processes.append(p)
    
    # 添加任务
    for i in range(10):
        task_queue.put(f"Task {i}")
    
    # 收集结果
    results = []
    for i in range(10):
        result = result_queue.get()
        results.append(result)
    
    # 结束工作进程
    for i in range(4):
        task_queue.put(None)
    
    for p in processes:
        p.join()
```

## 7. 配置化超时和重试

通过配置文件管理超时和重试策略：

```json
{
  "tool_settings": {
    "heartbeat_interval": 30,
    "max_retries": 5,
    "retry_delay": 2,
    "timeout": 300,
    "memory_limit_mb": 500
  },
  "persistence": {
    "enable_heartbeat": true,
    "enable_auto_recovery": true,
    "enable_resource_monitoring": true
  }
}
```

## 8. 使用系统级服务

在生产环境中，将工具作为系统服务运行：

### Windows服务示例 (使用NSSM)
```batch
nssm install WordDocTools "C:\Python\python.exe"
nssm set WordDocTools AppDirectory "C:\word-docx-tools"
nssm set WordDocTools AppParameters "main.py"
nssm set WordDocTools Start SERVICE_AUTO_START
```

### Linux systemd服务示例
```ini
[Unit]
Description=Word Document Tools
After=network.target

[Service]
Type=simple
User=tools-user
WorkingDirectory=/opt/word-docx-tools
ExecStart=/opt/venv/bin/python main.py
Restart=always
RestartSec=10

[Install]
WantedBy=multi-user.target
```

通过实施这些策略，可以确保工具的持久化执行，避免因调用过多或长时间运行而被系统终止。