/**
 * DeepSeek API 调用模块（优化版）
 * 文件名: DeepSeekAPI.gs
 * 支持标准调用和 Function Calling
 */

/**
 * DeepSeek API 配置
 * 请在脚本属性中设置 DEEPSEEK_API_KEY
 */
function getDeepSeekConfig() {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  return {
    apiKey: scriptProperties.getProperty('DEEPSEEK_API_KEY') || '',
    baseUrl: 'https://api.deepseek.com/v1/chat/completions',
    model: 'deepseek-chat',
    maxTokens: 4000,
    temperature: 0.7
  };
}

/**
 * 标准 API 调用（保留原有功能）
 * @param {string} prompt - 提示词
 * @param {number} maxTokens - 最大 token 数
 * @param {number} temperature - 温度参数
 * @returns {string} AI 响应
 */
function callDeepSeekAPI(prompt, maxTokens = null, temperature = null) {
  try {
    const config = getDeepSeekConfig();
    
    if (!config.apiKey) {
      return '错误: 未设置 API Key。请在脚本属性中设置 DEEPSEEK_API_KEY';
    }
    
    const payload = {
      model: config.model,
      messages: [
        {
          role: 'user',
          content: prompt
        }
      ],
      max_tokens: maxTokens || config.maxTokens,
      temperature: temperature !== null ? temperature : config.temperature
    };
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': 'Bearer ' + config.apiKey
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(config.baseUrl, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode !== 200) {
      Logger.log('API 错误响应: ' + responseText);
      return '错误: API 返回状态码 ' + responseCode + '\n' + responseText;
    }
    
    const jsonResponse = JSON.parse(responseText);
    
    if (jsonResponse.choices && jsonResponse.choices.length > 0) {
      return jsonResponse.choices[0].message.content;
    } else {
      return '错误: API 响应格式异常';
    }
    
  } catch (error) {
    Logger.log('调用 API 时出错: ' + error.message);
    return '错误: ' + error.message;
  }
}

/**
 * 支持 Function Calling 的 API 调用（新增）
 * @param {Array} messages - 消息数组
 * @param {Array} tools - 工具定义数组
 * @param {number} maxTokens - 最大 token 数
 * @param {number} temperature - 温度参数
 * @returns {Object} 包含响应和工具调用的对象
 */
function callDeepSeekAPIWithTools(messages, tools = null, maxTokens = null, temperature = null) {
  try {
    const config = getDeepSeekConfig();
    
    if (!config.apiKey) {
      return {
        error: '未设置 API Key。请在脚本属性中设置 DEEPSEEK_API_KEY'
      };
    }
    
    // 构建请求体
    const payload = {
      model: config.model,
      messages: messages,
      max_tokens: maxTokens || config.maxTokens,
      temperature: temperature !== null ? temperature : config.temperature
    };
    
    // 如果提供了工具定义,添加到请求中
    if (tools && tools.length > 0) {
      payload.tools = tools;
      payload.tool_choice = 'auto'; // 让模型自动决定是否使用工具
    }
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': 'Bearer ' + config.apiKey
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(config.baseUrl, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode !== 200) {
      Logger.log('API 错误响应: ' + responseText);
      return {
        error: 'API 返回状态码 ' + responseCode + ': ' + responseText
      };
    }
    
    const jsonResponse = JSON.parse(responseText);
    
    if (!jsonResponse.choices || jsonResponse.choices.length === 0) {
      return {
        error: 'API 响应格式异常'
      };
    }
    
    const choice = jsonResponse.choices[0];
    const message = choice.message;
    
    // 构建返回对象
    const result = {
      content: message.content || null,
      role: message.role || 'assistant'
    };
    
    // 如果有工具调用,添加到结果中
    if (message.tool_calls && message.tool_calls.length > 0) {
      result.tool_calls = message.tool_calls;
    }
    
    // 添加使用统计信息
    if (jsonResponse.usage) {
      result.usage = jsonResponse.usage;
    }
    
    return result;
    
  } catch (error) {
    Logger.log('调用 API 时出错: ' + error.message);
    return {
      error: error.message
    };
  }
}

/**
 * 测试 API 连接
 * @returns {string} 测试结果
 */
function testDeepSeekConnection() {
  try {
    const result = callDeepSeekAPI('你好,请用一句话介绍你自己。', 100);
    
    if (result.startsWith('错误:')) {
      return result;
    }
    
    return '连接成功！AI 回复: ' + result;
    
  } catch (error) {
    return '错误: ' + error.message;
  }
}

/**
 * 测试 Function Calling
 * @returns {Object} 测试结果
 */
function testFunctionCalling() {
  try {
    const messages = [
      {
        role: 'user',
        content: '当前工作簿有哪些工作表？'
      }
    ];
    
    const tools = [
      {
        type: 'function',
        function: {
          name: 'list_sheets',
          description: '获取当前工作簿中所有工作表的列表',
          parameters: {
            type: 'object',
            properties: {}
          }
        }
      }
    ];
    
    const result = callDeepSeekAPIWithTools(messages, tools);
    
    if (result.error) {
      return {
        success: false,
        error: result.error
      };
    }
    
    return {
      success: true,
      hasToolCalls: !!(result.tool_calls && result.tool_calls.length > 0),
      content: result.content,
      toolCalls: result.tool_calls
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 设置 API Key（辅助函数）
 * @param {string} apiKey - API 密钥
 */
function setDeepSeekAPIKey(apiKey) {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('DEEPSEEK_API_KEY', apiKey);
  return '✅ API Key 已保存';
}

/**
 * 获取 API Key（用于验证）
 * @returns {string} API Key（部分隐藏）
 */
function getDeepSeekAPIKey() {
  const config = getDeepSeekConfig();
  if (!config.apiKey) {
    return '❌ 未设置 API Key';
  }
  
  // 只显示前 8 位和后 4 位
  const key = config.apiKey;
  if (key.length > 12) {
    return key.substring(0, 8) + '...' + key.substring(key.length - 4);
  }
  return '***';
}

// ========== 新增功能：API Key 配置界面 ==========

/**
 * 显示 API Key 配置对话框（新增）
 * 提供用户友好的界面来输入和保存 API Key
 */
function showAPIKeyConfigDialog() {
  const html = HtmlService.createHtmlOutput(getConfigDialogHTML())
    .setWidth(500)
    .setHeight(350);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'DeepSeek API 配置');
}

/**
 * 生成配置对话框的 HTML（新增）
 * @returns {string} HTML 内容
 */
function getConfigDialogHTML() {
  const currentKey = getDeepSeekAPIKey();
  
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: Arial, sans-serif;
            padding: 20px;
            margin: 0;
          }
          .container {
            max-width: 460px;
          }
          .form-group {
            margin-bottom: 20px;
          }
          label {
            display: block;
            margin-bottom: 8px;
            font-weight: bold;
            color: #333;
          }
          input[type="text"] {
            width: 100%;
            padding: 10px;
            border: 1px solid #ddd;
            border-radius: 4px;
            box-sizing: border-box;
            font-size: 14px;
          }
          .current-key {
            background-color: #f5f5f5;
            padding: 10px;
            border-radius: 4px;
            margin-bottom: 15px;
            font-family: monospace;
            color: #666;
          }
          .button-group {
            display: flex;
            gap: 10px;
            margin-top: 20px;
          }
          button {
            flex: 1;
            padding: 12px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 14px;
            font-weight: bold;
          }
          .btn-primary {
            background-color: #4285f4;
            color: white;
          }
          .btn-primary:hover {
            background-color: #3367d6;
          }
          .btn-secondary {
            background-color: #f1f3f4;
            color: #5f6368;
          }
          .btn-secondary:hover {
            background-color: #e8eaed;
          }
          .btn-test {
            background-color: #34a853;
            color: white;
          }
          .btn-test:hover {
            background-color: #2d9048;
          }
          .message {
            padding: 12px;
            border-radius: 4px;
            margin-top: 15px;
            display: none;
          }
          .message.success {
            background-color: #d4edda;
            color: #155724;
            border: 1px solid #c3e6cb;
          }
          .message.error {
            background-color: #f8d7da;
            color: #721c24;
            border: 1px solid #f5c6cb;
          }
          .help-text {
            font-size: 12px;
            color: #666;
            margin-top: 5px;
          }
          .loading {
            display: none;
            text-align: center;
            padding: 10px;
            color: #666;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="form-group">
            <label>当前 API Key:</label>
            <div class="current-key">${currentKey}</div>
          </div>
          
          <div class="form-group">
            <label for="apiKey">输入新的 API Key:</label>
            <input type="text" id="apiKey" placeholder="sk-xxxxxxxxxxxxxxxxxxxxxxxx">
            <div class="help-text">
              请输入您的 DeepSeek API Key，可在 
              <a href="https://platform.deepseek.com/api_keys" target="_blank">DeepSeek 平台</a> 获取
            </div>
          </div>
          
          <div class="button-group">
            <button class="btn-primary" onclick="saveAPIKey()">保存</button>
            <button class="btn-test" onclick="testConnection()">测试连接</button>
            <button class="btn-secondary" onclick="google.script.host.close()">关闭</button>
          </div>
          
          <div id="message" class="message"></div>
          <div id="loading" class="loading">正在处理...</div>
        </div>
        
        <script>
          function saveAPIKey() {
            const apiKey = document.getElementById('apiKey').value.trim();
            const messageDiv = document.getElementById('message');
            const loadingDiv = document.getElementById('loading');
            
            if (!apiKey) {
              showMessage('请输入 API Key', 'error');
              return;
            }
            
            loadingDiv.style.display = 'block';
            messageDiv.style.display = 'none';
            
            google.script.run
              .withSuccessHandler(function(result) {
                loadingDiv.style.display = 'none';
                showMessage(result, 'success');
                document.getElementById('apiKey').value = '';
                
                // 更新当前 Key 显示
                setTimeout(function() {
                  google.script.run
                    .withSuccessHandler(function(key) {
                      document.querySelector('.current-key').textContent = key;
                    })
                    .getDeepSeekAPIKey();
                }, 500);
              })
              .withFailureHandler(function(error) {
                loadingDiv.style.display = 'none';
                showMessage('保存失败: ' + error.message, 'error');
              })
              .setDeepSeekAPIKey(apiKey);
          }
          
          function testConnection() {
            const messageDiv = document.getElementById('message');
            const loadingDiv = document.getElementById('loading');
            
            loadingDiv.style.display = 'block';
            messageDiv.style.display = 'none';
            
            google.script.run
              .withSuccessHandler(function(result) {
                loadingDiv.style.display = 'none';
                if (result.startsWith('错误:')) {
                  showMessage(result, 'error');
                } else {
                  showMessage(result, 'success');
                }
              })
              .withFailureHandler(function(error) {
                loadingDiv.style.display = 'none';
                showMessage('测试失败: ' + error.message, 'error');
              })
              .testDeepSeekConnection();
          }
          
          function showMessage(text, type) {
            const messageDiv = document.getElementById('message');
            messageDiv.textContent = text;
            messageDiv.className = 'message ' + type;
            messageDiv.style.display = 'block';
          }
          
          // 按 Enter 键保存
          document.getElementById('apiKey').addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
              saveAPIKey();
            }
          });
        </script>
      </body>
    </html>
  `;
}

/**
 * 在菜单中添加配置选项（新增）
 * 需要在主代码中的 onOpen() 函数调用
 */
//function addAPIConfigMenu() {
//  const ui = SpreadsheetApp.getUi();
//  ui.createMenu('DeepSeek API')
//    .addItem('配置 API Key', 'showAPIKeyConfigDialog')
//    .addItem('测试连接', 'testDeepSeekConnection')
//    .addItem('查看当前 Key', 'showCurrentAPIKey')
//    .addToUi();
//}

/**
 * 显示当前 API Key（新增）
 */
function showCurrentAPIKey() {
  const key = getDeepSeekAPIKey();
  const ui = SpreadsheetApp.getUi();
  ui.alert('当前 API Key', key, ui.ButtonSet.OK);
}

// ========== 原有功能保持不变 ==========

/**
 * 批量调用 API（用于处理多个请求）
 * @param {Array} prompts - 提示词数组
 * @returns {Array} 响应数组
 */
function batchCallDeepSeekAPI(prompts) {
  const results = [];
  
  prompts.forEach((prompt, index) => {
    try {
      const response = callDeepSeekAPI(prompt);
      results.push({
        index: index,
        success: !response.startsWith('错误:'),
        response: response
      });
      
      // 添加延迟避免频率限制
      if (index < prompts.length - 1) {
        Utilities.sleep(1000); // 1 秒延迟
      }
      
    } catch (error) {
      results.push({
        index: index,
        success: false,
        error: error.message
      });
    }
  });
  
  return results;
}

/**
 * 流式响应处理（模拟,GAS 不支持真正的流式）
 * @param {string} prompt - 提示词
 * @param {function} callback - 回调函数
 */
function streamDeepSeekAPI(prompt, callback) {
  // Google Apps Script 不支持真正的流式响应
  // 这里提供一个模拟实现
  
  try {
    const fullResponse = callDeepSeekAPI(prompt);
    
    if (fullResponse.startsWith('错误:')) {
      callback({ error: fullResponse });
      return;
    }
    
    // 模拟分段发送
    const chunkSize = 50;
    for (let i = 0; i < fullResponse.length; i += chunkSize) {
      const chunk = fullResponse.substring(i, i + chunkSize);
      callback({ chunk: chunk, done: i + chunkSize >= fullResponse.length });
      
      // 模拟延迟
      Utilities.sleep(100);
    }
    
  } catch (error) {
    callback({ error: error.message });
  }
}

/**
 * 估算 token 数量（粗略估计）
 * @param {string} text - 文本内容
 * @returns {number} 估计的 token 数
 */
function estimateTokenCount(text) {
  // 粗略估计:
  // - 英文: 约 4 字符 = 1 token
  // - 中文: 约 1.5 字符 = 1 token
  
  const chineseChars = (text.match(/[\u4e00-\u9fa5]/g) || []).length;
  const otherChars = text.length - chineseChars;
  
  const chineseTokens = chineseChars / 1.5;
  const otherTokens = otherChars / 4;
  
  return Math.ceil(chineseTokens + otherTokens);
}

/**
 * 检查 API 配额使用情况（需要额外的 API 支持）
 * @returns {Object} 配额信息
 */
function checkAPIQuota() {
  // 注意: DeepSeek API 可能不提供配额查询端点
  // 这里提供一个框架,具体实现需要查阅 API 文档
  
  return {
    available: true,
    message: 'DeepSeek API 配额查询功能暂未实现,请查看官方文档'
  };
}

