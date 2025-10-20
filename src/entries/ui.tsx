import React, { useEffect, useState, useCallback } from 'react';
import ReactDOM from 'react-dom/client';
import { initView } from 'dingtalk-docs-cool-app';

interface Sheet {
  id: string;
  name: string;
  desc: string;
  fieldsCount: number;
}

type MergeMode = 'create' | 'overwrite' | 'append';

// 配置接口
interface SavedConfig {
  selectedSheets: string[];
  mergeSheetName: string;
  mergeMode: MergeMode;
  targetSheetId: string;
  manualTargetSheetName: string;
  timestamp: number;
}

function App() {
  const [currentSheetName, setCurrentSheetName] = useState<string>('');
  const [sheets, setSheets] = useState<Sheet[]>([]);
  const [selectedSheets, setSelectedSheets] = useState<string[]>([]);
  const [mergeSheetName, setMergeSheetName] = useState<string>('');
  const [mergeMode, setMergeMode] = useState<MergeMode>('create');
  const [targetSheetId, setTargetSheetId] = useState<string>('');
  const [manualTargetSheetName, setManualTargetSheetName] = useState<string>('');
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [message, setMessage] = useState<string>('');
  const [hasSavedConfig, setHasSavedConfig] = useState<boolean>(false);
  const [manualSheetNameValid, setManualSheetNameValid] = useState<boolean | null>(null);

  // 保存配置到localStorage
  const saveConfig = useCallback(() => {
    const config: SavedConfig = {
      selectedSheets,
      mergeSheetName,
      mergeMode,
      targetSheetId,
      manualTargetSheetName,
      timestamp: Date.now()
    };
    console.log('保存配置:', config);
    localStorage.setItem('mergeSheetsConfig', JSON.stringify(config));
    setHasSavedConfig(true);
    setMessage('配置已保存');
    console.log('配置保存完成');
  }, [selectedSheets, mergeSheetName, mergeMode, targetSheetId, manualTargetSheetName]);

  // 从localStorage恢复配置
  const restoreConfig = useCallback(() => {
    try {
      const savedConfig = localStorage.getItem('mergeSheetsConfig');
      if (savedConfig) {
        const config: SavedConfig = JSON.parse(savedConfig);
        
        // 检查配置是否过期（24小时）
        const isExpired = Date.now() - config.timestamp > 24 * 60 * 60 * 1000;
        if (isExpired) {
          localStorage.removeItem('mergeSheetsConfig');
          setHasSavedConfig(false);
          setMessage('保存的配置已过期，已清除');
          return;
        }

        // 验证保存的数据表ID是否仍然存在
        const validSheetIds = config.selectedSheets.filter(id => 
          sheets.some(sheet => sheet.id === id)
        );

        setSelectedSheets(validSheetIds);
        setMergeSheetName(config.mergeSheetName);
        setMergeMode(config.mergeMode);
        
        // 验证目标表ID是否仍然存在
        if (config.targetSheetId && sheets.some(sheet => sheet.id === config.targetSheetId)) {
          setTargetSheetId(config.targetSheetId);
        } else {
          setTargetSheetId('');
        }
        
        // 恢复手动输入的数据表名
        setManualTargetSheetName(config.manualTargetSheetName || '');

        setMessage(`已恢复上次配置（${validSheetIds.length}个数据表）`);
      }
    } catch (error) {
      console.error('恢复配置失败:', error);
      setMessage('恢复配置失败');
    }
  }, [sheets]);

  // 清除保存的配置
  const clearConfig = useCallback(() => {
    localStorage.removeItem('mergeSheetsConfig');
    setHasSavedConfig(false);
    setMessage('已清除保存的配置');
  }, []);

  // 检查是否有保存的配置
  const checkSavedConfig = useCallback(() => {
    const savedConfig = localStorage.getItem('mergeSheetsConfig');
    if (savedConfig) {
      try {
        const config: SavedConfig = JSON.parse(savedConfig);
        const isExpired = Date.now() - config.timestamp > 24 * 60 * 60 * 1000;
        setHasSavedConfig(!isExpired);
      } catch (error) {
        setHasSavedConfig(false);
      }
    } else {
      setHasSavedConfig(false);
    }
  }, []);

  const getCurrentSheetName = useCallback(async () => {
    try {
      // 检查脚本是否准备就绪
      if (!Dingdocs?.script) {
        console.log('脚本尚未准备就绪，等待中...');
        setCurrentSheetName('脚本加载中...');
        return;
      }
      
      const name = await Dingdocs.script.run('getActiveSheetName');
      setCurrentSheetName(name || '未找到数据表');
    } catch (error) {
      console.error('获取数据表名称失败:', error);
      if (error.message?.includes('ScriptSandboxNotReady')) {
        setCurrentSheetName('脚本加载中...');
        // 如果脚本未准备就绪，延迟重试
        setTimeout(() => {
          getCurrentSheetName();
        }, 1000);
      } else {
        setCurrentSheetName('获取失败');
      }
    }
  }, []);

  const getAllSheets = useCallback(async () => {
    try {
      if (!Dingdocs?.script) {
        console.log('脚本尚未准备就绪');
        return;
      }
      
      const sheetsData = await Dingdocs.script.run('getAllSheets');
      setSheets(sheetsData || []);
      
      // 获取数据表列表后检查保存的配置
      checkSavedConfig();
    } catch (error) {
      console.error('获取数据表列表失败:', error);
      setMessage('获取数据表列表失败');
    }
  }, [checkSavedConfig]);

  const handleSheetToggle = (sheetId: string) => {
    setSelectedSheets(prev => 
      prev.includes(sheetId) 
        ? prev.filter(id => id !== sheetId)
        : [...prev, sheetId]
    );
  };

  const handleSelectAll = () => {
    if (selectedSheets.length === sheets.length) {
      // 如果全部选中，则取消全选
      setSelectedSheets([]);
    } else {
      // 否则全选
      setSelectedSheets(sheets.map(sheet => sheet.id));
    }
  };

  const handleMergeModeChange = (mode: MergeMode) => {
    setMergeMode(mode);
    if (mode === 'create') {
      setTargetSheetId('');
      setManualTargetSheetName('');
    }
  };

  // 根据数据表名称查找数据表ID
  const findSheetIdByName = useCallback((sheetName: string) => {
    const trimmedName = sheetName.trim();
    // 首先尝试精确匹配（忽略前后空格）
    let sheet = sheets.find(s => s.name === trimmedName);
    
    // 如果精确匹配失败，尝试忽略大小写的匹配
    if (!sheet) {
      sheet = sheets.find(s => s.name.toLowerCase() === trimmedName.toLowerCase());
    }
    
    return sheet ? sheet.id : null;
  }, [sheets]);

  const handleMergeSheets = async () => {
    if (selectedSheets.length < 1) {
      setMessage('请至少选择1个数据表');
      return;
    }
    
    if (mergeMode === 'create' && !mergeSheetName.trim()) {
      setMessage('请输入合并后的数据表名称');
      return;
    }

    // 验证目标数据表（覆盖和追加模式）
    if (mergeMode === 'overwrite' || mergeMode === 'append') {
      // 优先使用手动输入的数据表名
      if (manualTargetSheetName.trim()) {
        const foundSheetId = findSheetIdByName(manualTargetSheetName.trim());
        if (!foundSheetId) {
          setMessage(`未找到名称为 "${manualTargetSheetName.trim()}" 的数据表`);
          return;
        }
      } else if (!targetSheetId) {
        setMessage('请选择目标数据表或输入数据表名称');
        return;
      }
    }

    setIsLoading(true);
    setMessage('');

    try {
      if (!Dingdocs?.script) {
        throw new Error('脚本尚未准备就绪');
      }

      // 确定最终的目标数据表ID
      let finalTargetSheetId = targetSheetId;
      if (mergeMode === 'overwrite' || mergeMode === 'append') {
        if (manualTargetSheetName.trim()) {
          const foundSheetId = findSheetIdByName(manualTargetSheetName.trim());
          if (foundSheetId) {
            finalTargetSheetId = foundSheetId;
          }
        }
      }

      const result = await Dingdocs.script.run('mergeSheets', {
        selectedSheetIds: selectedSheets,
        mergeSheetName: mergeSheetName.trim(),
        mergeMode: mergeMode,
        targetSheetId: finalTargetSheetId
      });

      let successMessage = '';
      if (mergeMode === 'create') {
        if (selectedSheets.length === 1) {
          successMessage = `复制成功！新数据表 "${result.name}" 已创建，包含 ${result.totalRecords} 条记录`;
        } else {
          successMessage = `合并成功！新数据表 "${result.name}" 已创建，包含 ${result.totalRecords} 条记录`;
        }
      } else if (mergeMode === 'overwrite') {
        if (selectedSheets.length === 1) {
          successMessage = `覆盖成功！数据表 "${result.name}" 已更新，包含 ${result.totalRecords} 条记录`;
        } else {
          successMessage = `覆盖成功！数据表 "${result.name}" 已更新，包含 ${result.totalRecords} 条记录`;
        }
      } else {
        if (selectedSheets.length === 1) {
          successMessage = `追加成功！数据表 "${result.name}" 已添加 ${result.addedRecords} 条记录`;
        } else {
          successMessage = `追加成功！数据表 "${result.name}" 已添加 ${result.addedRecords} 条记录`;
        }
      }

      setMessage(successMessage);
      
      // 自动保存当前配置
      saveConfig();
      
      // 延迟刷新数据表列表，避免覆盖消息
      setTimeout(async () => {
        await getAllSheets();
      }, 1000);
    } catch (error: any) {
      console.error('合并数据表失败:', error);
      setMessage(`合并失败: ${error.message}`);
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    initView({
      onReady: () => {
        console.log('init ui completed!');
        // 初始化完成后尝试获取当前数据表名称和所有数据表列表
        getCurrentSheetName();
        getAllSheets();
      },
      onError: (e) => {
        console.log(e);
      },
    });
  }, [getCurrentSheetName, getAllSheets]);

  return (
    <div style={{ padding: '16px', fontFamily: 'Arial, sans-serif' }}>
      <h1>合并数据表</h1>
      
      {/* 当前数据表信息 */}
      <div style={{ marginBottom: '20px', padding: '10px', backgroundColor: '#f5f5f5', borderRadius: '4px' }}>
        <p><strong>当前数据表：</strong>{currentSheetName}</p>
        <div style={{ marginTop: '10px' }}>
          <button onClick={getCurrentSheetName} style={{ marginRight: '10px' }}>刷新</button>
          <button onClick={getAllSheets} style={{ marginRight: '10px' }}>刷新数据表列表</button>
          {hasSavedConfig && (
            <>
              <button onClick={restoreConfig} style={{ marginRight: '10px', backgroundColor: '#52c41a', color: 'white', border: 'none', padding: '4px 8px', borderRadius: '4px' }}>
                恢复上次配置
              </button>
              <button onClick={clearConfig} style={{ backgroundColor: '#ff4d4f', color: 'white', border: 'none', padding: '4px 8px', borderRadius: '4px' }}>
                清除配置
              </button>
            </>
          )}
        </div>
      </div>

      {/* 数据表选择 */}
      <div style={{ marginBottom: '20px' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '10px' }}>
          <h3>选择要处理的数据表：</h3>
          <button
            onClick={handleSelectAll}
            style={{
              padding: '4px 8px',
              fontSize: '12px',
              backgroundColor: '#f0f0f0',
              border: '1px solid #ddd',
              borderRadius: '4px',
              cursor: 'pointer'
            }}
          >
            {selectedSheets.length === sheets.length ? '取消全选' : '全选'}
          </button>
        </div>
        <div style={{ maxHeight: '200px', overflowY: 'auto', border: '1px solid #ddd', padding: '10px' }}>
          {sheets.length === 0 ? (
            <p>暂无数据表</p>
          ) : (
            sheets.map(sheet => (
              <label key={sheet.id} style={{ display: 'block', marginBottom: '8px', cursor: 'pointer' }}>
                <input
                  type="checkbox"
                  checked={selectedSheets.includes(sheet.id)}
                  onChange={() => handleSheetToggle(sheet.id)}
                  style={{ marginRight: '8px' }}
                />
                <span>
                  {sheet.name} 
                  <span style={{ color: '#666', fontSize: '12px' }}>
                    （{sheet.fieldsCount}个字段）
                  </span>
                </span>
              </label>
            ))
          )}
        </div>
        <p style={{ fontSize: '12px', color: '#666', marginTop: '5px' }}>
          已选择 {selectedSheets.length} 个数据表
        </p>
      </div>

      {/* 合并设置 */}
      <div style={{ marginBottom: '20px' }}>
        <h3>合并设置：</h3>
        
        {/* 合并模式选择 */}
        <div style={{ marginBottom: '15px' }}>
          <label style={{ display: 'block', marginBottom: '8px', fontWeight: 'bold' }}>
            合并模式：
          </label>
          <div style={{ display: 'flex', gap: '10px', flexWrap: 'wrap' }}>
            <label style={{ display: 'flex', alignItems: 'center', cursor: 'pointer' }}>
              <input
                type="radio"
                name="mergeMode"
                value="create"
                checked={mergeMode === 'create'}
                onChange={() => handleMergeModeChange('create')}
                style={{ marginRight: '5px' }}
              />
              创建新表
            </label>
            <label style={{ display: 'flex', alignItems: 'center', cursor: 'pointer' }}>
              <input
                type="radio"
                name="mergeMode"
                value="overwrite"
                checked={mergeMode === 'overwrite'}
                onChange={() => handleMergeModeChange('overwrite')}
                style={{ marginRight: '5px' }}
              />
              覆盖已有表
            </label>
            <label style={{ display: 'flex', alignItems: 'center', cursor: 'pointer' }}>
              <input
                type="radio"
                name="mergeMode"
                value="append"
                checked={mergeMode === 'append'}
                onChange={() => handleMergeModeChange('append')}
                style={{ marginRight: '5px' }}
              />
              追加写入
            </label>
          </div>
        </div>

        {/* 根据模式显示不同的输入 */}
        {mergeMode === 'create' && (
          <div>
            <label style={{ display: 'block', marginBottom: '5px' }}>
              新数据表名称：
            </label>
            <input
              type="text"
              value={mergeSheetName}
              onChange={(e) => setMergeSheetName(e.target.value)}
              placeholder="请输入合并后的数据表名称"
              style={{ 
                width: '100%', 
                padding: '8px', 
                border: '1px solid #ddd', 
                borderRadius: '4px',
                boxSizing: 'border-box'
              }}
            />
          </div>
        )}

        {(mergeMode === 'overwrite' || mergeMode === 'append') && (
          <div>
            <label style={{ display: 'block', marginBottom: '5px' }}>
              目标数据表：
            </label>
            
            {/* 手动输入数据表名 */}
            <div style={{ marginBottom: '10px' }}>
              <input
                type="text"
                value={manualTargetSheetName}
                onChange={(e) => {
                  setManualTargetSheetName(e.target.value);
                  // 如果手动输入了内容，清空下拉选择
                  if (e.target.value.trim()) {
                    setTargetSheetId('');
                  }
                }}
                placeholder="手动输入数据表名称（优先级高于下拉选择）"
                style={{ 
                  width: '100%', 
                  padding: '8px', 
                  border: '1px solid #ddd', 
                  borderRadius: '4px',
                  boxSizing: 'border-box',
                  backgroundColor: manualTargetSheetName.trim() ? '#f0f8ff' : 'white'
                }}
              />
              {manualTargetSheetName.trim() && (
                <div style={{ fontSize: '12px', color: '#666', marginTop: '2px' }}>
                  手动输入模式：将使用 "{manualTargetSheetName.trim()}" 作为目标数据表
                </div>
              )}
            </div>
            
            {/* 下拉选择数据表 */}
            <div style={{ opacity: manualTargetSheetName.trim() ? 0.5 : 1 }}>
              <select
                value={targetSheetId}
                onChange={(e) => {
                  setTargetSheetId(e.target.value);
                  // 如果选择了下拉选项，清空手动输入
                  if (e.target.value) {
                    setManualTargetSheetName('');
                  }
                }}
                disabled={!!manualTargetSheetName.trim()}
                style={{ 
                  width: '100%', 
                  padding: '8px', 
                  border: '1px solid #ddd', 
                  borderRadius: '4px',
                  boxSizing: 'border-box'
                }}
              >
                <option value="">请选择目标数据表</option>
                {sheets.map(sheet => (
                  <option key={sheet.id} value={sheet.id}>
                    {sheet.name} ({sheet.fieldsCount}个字段)
                  </option>
                ))}
              </select>
            </div>
          </div>
        )}
      </div>

      {/* 操作按钮 */}
      <div style={{ marginBottom: '20px' }}>
        <button
          onClick={handleMergeSheets}
          disabled={
            isLoading || 
            selectedSheets.length < 1 || 
            (mergeMode === 'create' && !mergeSheetName.trim()) ||
            ((mergeMode === 'overwrite' || mergeMode === 'append') && !targetSheetId && !manualTargetSheetName.trim())
          }
          style={{
            padding: '10px 20px',
            backgroundColor: (
              selectedSheets.length >= 1 && 
              ((mergeMode === 'create' && mergeSheetName.trim()) || 
               ((mergeMode === 'overwrite' || mergeMode === 'append') && (targetSheetId || manualTargetSheetName.trim())))
            ) ? '#1890ff' : '#ccc',
            color: 'white',
            border: 'none',
            borderRadius: '4px',
            cursor: (
              selectedSheets.length >= 1 && 
              ((mergeMode === 'create' && mergeSheetName.trim()) || 
               ((mergeMode === 'overwrite' || mergeMode === 'append') && (targetSheetId || manualTargetSheetName.trim())))
            ) ? 'pointer' : 'not-allowed'
          }}
        >
          {isLoading ? '处理中...' : 
           mergeMode === 'create' ? '创建新表' :
           mergeMode === 'overwrite' ? '覆盖数据表' : '追加数据'}
        </button>
        
        <button
          onClick={saveConfig}
          disabled={isLoading}
          style={{
            marginLeft: '10px',
            padding: '10px 20px',
            backgroundColor: '#52c41a',
            color: 'white',
            border: 'none',
            borderRadius: '4px',
            cursor: isLoading ? 'not-allowed' : 'pointer'
          }}
        >
          保存当前配置
        </button>
      </div>

      {/* 消息显示 */}
      {message && (
        <div style={{
          padding: '10px',
          backgroundColor: message.includes('成功') ? '#f6ffed' : '#fff2f0',
          border: `1px solid ${message.includes('成功') ? '#b7eb8f' : '#ffccc7'}`,
          borderRadius: '4px',
          color: message.includes('成功') ? '#52c41a' : '#ff4d4f'
        }}>
          {message}
        </div>
      )}

      {/* 功能说明 */}
      <div style={{ marginTop: '20px', fontSize: '12px', color: '#666' }}>
        <h4>功能说明：</h4>
        <ul>
          <li><strong>创建新表</strong>：将选中的数据表合并到全新的数据表</li>
          <li><strong>覆盖已有表</strong>：清空目标表，用选中的数据替换</li>
          <li><strong>追加写入</strong>：在目标表末尾添加选中的数据</li>
          <li><strong>智能字段映射</strong>：自动识别公共字段，独有字段重命名</li>
          <li><strong>选项自动合并</strong>：单选字段选项自动补充完整</li>
          <li><strong>来源追踪</strong>：记录每条数据来自哪个数据表</li>
          <li><strong>支持单表处理</strong>：可以选择1个或多个数据表</li>
          <li><strong>配置保存</strong>：自动保存上次执行配置，支持一键恢复</li>
        </ul>
      </div>
    </div>
  )
}

const root = ReactDOM.createRoot(document.getElementById('ui')!);
root.render(
  <React.StrictMode>
    <App/>
  </React.StrictMode>
);