#!/usr/bin/env python3
"""MCP initialize と tools/list のみ。接続・認証・アプリ変更は行わない。"""
import json, os, queue, subprocess, threading

command = os.environ.get('CANVAS_DNX', 'dnx')
p = subprocess.Popen([command, 'Microsoft.PowerApps.CanvasAuthoring.McpServer', '--yes'],
                     stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, text=True)
lines = queue.Queue()
def read_lines():
    for line in p.stdout:
        lines.put(line)
threading.Thread(target=read_lines, daemon=True).start()
def send(payload):
    p.stdin.write(json.dumps(payload) + '\n')
    p.stdin.flush()
def response(request_id):
    import time
    deadline = time.monotonic() + 120
    while time.monotonic() < deadline:
        try: line = lines.get(timeout=min(1, max(.01, deadline-time.monotonic())))
        except queue.Empty:
            if p.poll() is not None: raise RuntimeError('MCP process exited before response')
            continue
        try: result = json.loads(line)
        except ValueError: continue
        if result.get('id') == request_id:
            if 'error' in result: raise RuntimeError('MCP returned a protocol error')
            return result['result']
    raise TimeoutError('MCP response timed out; check .NET and NuGet access')
try:
    send({'jsonrpc':'2.0','id':1,'method':'initialize','params':{'protocolVersion':'2024-11-05','capabilities':{},'clientInfo':{'name':'canvas-readonly-probe','version':'1.0'}}})
    result = response(1)
    print(json.dumps({'serverInfo':result.get('serverInfo'),'protocolVersion':result.get('protocolVersion')}, ensure_ascii=False))
    send({'jsonrpc':'2.0','method':'notifications/initialized'})
    send({'jsonrpc':'2.0','id':2,'method':'tools/list','params':{}})
    result = response(2)
    print(json.dumps({'tools':[t['name'] for t in result.get('tools',[])]}, ensure_ascii=False))
finally:
    p.terminate()
    try: p.wait(timeout=5)
    except subprocess.TimeoutExpired: p.kill(); p.wait()
