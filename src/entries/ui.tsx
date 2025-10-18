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

function App() {
  const [currentSheetName, setCurrentSheetName] = useState<string>('');
  const [sheets, setSheets] = useState<Sheet[]>([]);
  const [selectedSheets, setSelectedSheets] = useState<string[]>([]);
  const [mergeSheetName, setMergeSheetName] = useState<string>('');
  const [mergeMode, setMergeMode] = useState<MergeMode>('create');
  const [targetSheetId, setTargetSheetId] = useState<string>('');
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [message, setMessage] = useState<string>('');

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
    } catch (error) {
      console.error('获取数据表列表失败:', error);
      setMessage('获取数据表列表失败');
    }
  }, []);

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
    }
  };

  const handleMergeSheets = async () => {
    if (selectedSheets.length < 2) {
      setMessage('请至少选择2个数据表进行合并');
      return;
    }
    
    if (mergeMode === 'create' && !mergeSheetName.trim()) {
      setMessage('请输入合并后的数据表名称');
      return;
    }

    if ((mergeMode === 'overwrite' || mergeMode === 'append') && !targetSheetId) {
      setMessage('请选择目标数据表');
      return;
    }

    setIsLoading(true);
    setMessage('');

    try {
      if (!Dingdocs?.script) {
        throw new Error('脚本尚未准备就绪');
      }

      const result = await Dingdocs.script.run('mergeSheets', {
        selectedSheetIds: selectedSheets,
        mergeSheetName: mergeSheetName.trim(),
        mergeMode: mergeMode,
        targetSheetId: targetSheetId
      });

      let successMessage = '';
      if (mergeMode === 'create') {
        successMessage = `合并成功！新数据表 "${result.name}" 已创建，包含 ${result.totalRecords} 条记录`;
      } else if (mergeMode === 'overwrite') {
        successMessage = `覆盖成功！数据表 "${result.name}" 已更新，包含 ${result.totalRecords} 条记录`;
      } else {
        successMessage = `追加成功！数据表 "${result.name}" 已添加 ${result.addedRecords} 条记录`;
      }

      setMessage(successMessage);
      setSelectedSheets([]);
      setMergeSheetName('');
      setTargetSheetId('');
      // 刷新数据表列表
      await getAllSheets();
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
        <button onClick={getCurrentSheetName} style={{ marginRight: '10px' }}>刷新</button>
        <button onClick={getAllSheets}>刷新数据表列表</button>
      </div>

      {/* 数据表选择 */}
      <div style={{ marginBottom: '20px' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '10px' }}>
          <h3>选择要合并的数据表（至少选择2个）：</h3>
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
            <select
              value={targetSheetId}
              onChange={(e) => setTargetSheetId(e.target.value)}
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
        )}
      </div>

      {/* 操作按钮 */}
      <div style={{ marginBottom: '20px' }}>
        <button
          onClick={handleMergeSheets}
          disabled={
            isLoading || 
            selectedSheets.length < 2 || 
            (mergeMode === 'create' && !mergeSheetName.trim()) ||
            ((mergeMode === 'overwrite' || mergeMode === 'append') && !targetSheetId)
          }
          style={{
            padding: '10px 20px',
            backgroundColor: (
              selectedSheets.length >= 2 && 
              ((mergeMode === 'create' && mergeSheetName.trim()) || 
               ((mergeMode === 'overwrite' || mergeMode === 'append') && targetSheetId))
            ) ? '#1890ff' : '#ccc',
            color: 'white',
            border: 'none',
            borderRadius: '4px',
            cursor: (
              selectedSheets.length >= 2 && 
              ((mergeMode === 'create' && mergeSheetName.trim()) || 
               ((mergeMode === 'overwrite' || mergeMode === 'append') && targetSheetId))
            ) ? 'pointer' : 'not-allowed'
          }}
        >
          {isLoading ? '合并中...' : 
           mergeMode === 'create' ? '创建新表' :
           mergeMode === 'overwrite' ? '覆盖数据表' : '追加数据'}
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
          <li><strong>创建新表</strong>：合并数据到全新的数据表</li>
          <li><strong>覆盖已有表</strong>：清空目标表，用合并数据替换</li>
          <li><strong>追加写入</strong>：在目标表末尾添加合并数据</li>
          <li><strong>智能字段映射</strong>：自动识别公共字段，放在前面</li>
          <li><strong>添加来源字段</strong>：记录每条数据来自哪个数据表</li>
          <li><strong>全选功能</strong>：一键选择所有数据表</li>
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