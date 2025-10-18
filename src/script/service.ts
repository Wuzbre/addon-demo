/*global DingdocsScript*/

/**
 * AI表格边栏插件服务层
 * 运行在 Web Worker 中，提供AI表格操作的核心功能
 */

// 获取当前激活的数据表
function getActiveSheet() {
  try {
    const base = DingdocsScript.base;
    const sheet = base.getActiveSheet();
    if (!sheet) {
      throw new Error('未找到激活的数据表');
    }
    return {
      id: sheet.getId(),
      name: sheet.getName(),
      desc: sheet.getDesc() || '',
      fieldsCount: sheet.getFields().length
    };
  } catch (error: any) {
    throw new Error(`获取激活数据表失败: ${error.message}`);
  }
}

// 获取所有数据表列表
function getAllSheets() {
  try {
    const base = DingdocsScript.base;
    const sheets = base.getSheets();
    return sheets.map((sheet: any) => ({
      id: sheet.getId(),
      name: sheet.getName(),
      desc: sheet.getDesc() || '',
      fieldsCount: sheet.getFields().length
    }));
  } catch (error: any) {
    throw new Error(`获取数据表列表失败: ${error.message}`);
  }
}

// 创建新的数据表
function createSheet(name: string) {
  try {
    if (!name || name.trim() === '') {
      throw new Error('数据表名称不能为空');
    }
    
    const base = DingdocsScript.base;
    // 创建带有基本字段的数据表
    const sheet = base.insertSheet(name.trim(), [
      { name: '标题', type: 'text' },
      { name: '状态', type: 'singleSelect' },
      { name: '创建时间', type: 'date' }
    ]);
    
    return {
      id: sheet.getId(),
      name: sheet.getName(),
      desc: sheet.getDesc() || '',
      fieldsCount: sheet.getFields().length
    };
  } catch (error: any) {
    throw new Error(`创建数据表失败: ${error.message}`);
  }
}

// 删除数据表
function deleteSheet(sheetId: string) {
  try {
    if (!sheetId) {
      throw new Error('数据表ID不能为空');
    }
    console.log('sheetId', sheetId);
    const base = DingdocsScript.base;
    base.deleteSheet(sheetId);
    return { success: true };
  } catch (error: any) {
    throw new Error(`删除数据表失败: ${error.message}`);
  }
}

// 获取数据表字段信息
function getSheetFields(sheetId?: string) {
  try {
    const base = DingdocsScript.base;
    let sheet;
    if (sheetId) {
      sheet = base.getSheet(sheetId);
    } else {
      sheet = base.getActiveSheet();
    }
    
    if (!sheet) {
      throw new Error('未找到指定的数据表');
    }
    
    const fields = sheet.getFields();
    return fields.map((field: any) => ({
      id: field.getId(),
      name: field.getName(),
      type: field.getType(),
      isPrimary: field.isPrimary?.() || false,
    }));
  } catch (error: any) {
    throw new Error(`获取字段信息失败: ${error.message}`);
  }
}

// 添加字段
function addField(name: string, type: string, sheetId?: string) {
  try {
    if (!name || name.trim() === '') {
      throw new Error('字段名称不能为空');
    }
    
    const base = DingdocsScript.base;
    let sheet;
    if (sheetId) {
      sheet = base.getSheet(sheetId);
    } else {
      sheet = base.getActiveSheet();
    }
    
    if (!sheet) {
      throw new Error('未找到指定的数据表');
    }
    
    const field = sheet.insertField({
      name: name.trim(),
      type: type as any
    });
    
    return {
      id: field.getId(),
      name: field.getName(),
      type: field.getType(),
      isPrimary: field.isPrimary?.() || false
    };
  } catch (error: any) {
    throw new Error(`添加字段失败: ${error.message}`);
  }
}

// 删除字段
function deleteField(fieldId: string, sheetId?: string) {
  try {
    if (!fieldId) {
      throw new Error('字段ID不能为空');
    }
    
    const base = DingdocsScript.base;
    let sheet;
    if (sheetId) {
      sheet = base.getSheet(sheetId);
    } else {
      sheet = base.getActiveSheet();
    }
    
    if (!sheet) {
      throw new Error('未找到指定的数据表');
    }
    
    // 检查是否为主键字段
    const field = sheet.getField(fieldId);
    if (field && field.isPrimary?.()) {
      throw new Error('不能删除主键字段');
    }
    
    sheet.deleteField(fieldId);
    return { success: true };
  } catch (error: any) {
    throw new Error(`删除字段失败: ${error.message}`);
  }
}

// 获取记录数据
async function getRecords(sheetId?: string, pageSize = 20) {
  try {
    const base = DingdocsScript.base;
    let sheet;
    if (sheetId) {
      sheet = base.getSheet(sheetId);
    } else {
      sheet = base.getActiveSheet();
    }
    
    if (!sheet) {
      throw new Error('未找到指定的数据表');
    }
    
    const result = await sheet.getRecordsAsync({ pageSize });
    
    return {
      records: result.records.map((record: any) => ({
        id: record.getId(),
        fields: record.getCellValues()
      })),
      hasMore: result.hasMore,
      cursor: result.cursor,
      total: result.records.length
    };
  } catch (error: any) {
    throw new Error(`获取记录失败: ${error.message}`);
  }
}

// 添加记录
async function addRecord(fields: Record<string, any>, sheetId?: string) {
  try {
    const base = DingdocsScript.base;
    let sheet;
    if (sheetId) {
      sheet = base.getSheet(sheetId);
    } else {
      sheet = base.getActiveSheet();
    }
    
    if (!sheet) {
      throw new Error('未找到指定的数据表');
    }
    
    const records = await sheet.insertRecordsAsync([{ fields }]);
    const record = records[0];
    
    return {
      id: record.getId(),
      fields: record.getCellValues()
    };
  } catch (error: any) {
    throw new Error(`添加记录失败: ${error.message}`);
  }
}

// 更新记录
async function updateRecord(recordId: string, fields: Record<string, any>, sheetId?: string) {
  try {
    if (!recordId) {
      throw new Error('记录ID不能为空');
    }
    
    const base = DingdocsScript.base;
    let sheet;
    if (sheetId) {
      sheet = base.getSheet(sheetId);
    } else {
      sheet = base.getActiveSheet();
    }
    
    if (!sheet) {
      throw new Error('未找到指定的数据表');
    }
    
    const records = await sheet.updateRecordsAsync([{ id: recordId, fields }]);
    const record = records[0];
    
    return {
      id: record.getId(),
      fields: record.getCellValues()
    };
  } catch (error: any) {
    throw new Error(`更新记录失败: ${error.message}`);
  }
}

// 删除记录
async function deleteRecord(recordId: string, sheetId?: string) {
  try {
    if (!recordId) {
      throw new Error('记录ID不能为空');
    }
    
    const base = DingdocsScript.base;
    let sheet;
    if (sheetId) {
      sheet = base.getSheet(sheetId);
    } else {
      sheet = base.getActiveSheet();
    }
    
    if (!sheet) {
      throw new Error('未找到指定的数据表');
    }
    
    await sheet.deleteRecordsAsync(recordId);
    return { success: true };
  } catch (error: any) {
    throw new Error(`删除记录失败: ${error.message}`);
  }
}

// 获取文档信息
function getDocumentInfo() {
  try {
    const base = DingdocsScript.base;
    const uuid = base.getDentryUuid();
    const sheets = base.getSheets();
    
    return {
      uuid,
      sheetsCount: sheets.length,
      currentSheet: base.getActiveSheet()?.getName() || '无'
    };
  } catch (error: any) {
    throw new Error(`获取文档信息失败: ${error.message}`);
  }
}

// 1.定义如何读写文档对象模型
function getActiveSheetName() {
  const sheet = DingdocsScript.base.getActiveSheet();
  return sheet?.getName();
}

// 合并数据表
async function mergeSheets(params: { 
  selectedSheetIds: string[]; 
  mergeSheetName: string; 
  mergeMode: 'create' | 'overwrite' | 'append';
  targetSheetId?: string;
}) {
  try {
    const { selectedSheetIds, mergeSheetName, mergeMode, targetSheetId } = params;
    
    if (!selectedSheetIds || selectedSheetIds.length < 2) {
      throw new Error('至少需要选择2个数据表进行合并');
    }

    const base = DingdocsScript.base;
    
    // 获取所有要合并的数据表
    const sheetsToMerge = selectedSheetIds.map(sheetId => {
      const sheet = base.getSheet(sheetId);
      if (!sheet) {
        throw new Error(`未找到ID为 ${sheetId} 的数据表`);
      }
      return sheet;
    });

    let targetSheet: any;
    let totalRecords = 0;
    let addedRecords = 0;

    if (mergeMode === 'create') {
      // 创建新表模式
      if (!mergeSheetName || mergeSheetName.trim() === '') {
        throw new Error('合并后的数据表名称不能为空');
      }

      const trimmedName = mergeSheetName.trim();
      
      // 检查名称是否重复
      const existingSheets = base.getSheets();
      if (existingSheets.some((sheet: any) => sheet.getName() === trimmedName)) {
        throw new Error(`数据表名称 "${trimmedName}" 已存在，请使用其他名称`);
      }

      // 分析字段映射
      const fieldMapping = analyzeFieldMapping(sheetsToMerge);
      
      // 准备字段配置，包含来源字段
      const allFieldConfigs = [
        ...fieldMapping.fields.map(field => ({
          name: field.name,
          type: field.type,
          property: field.property
        })),
        {
          name: '合并来源',
          type: 'text'
        }
      ];
      
      console.log('准备创建的字段:', allFieldConfigs.map(f => f.name));
      
      // 直接使用字段配置创建新表，避免空表+逐个创建的问题
      targetSheet = base.insertSheet(trimmedName, allFieldConfigs as any);
      
      // 检查创建后的字段
      const createdFields = targetSheet.getFields();
      console.log('创建新表后的字段数量:', createdFields.length);
      console.log('创建后的字段列表:', createdFields.map((f: any) => f.getName()));
      
      // 创建字段名映射表（字段名应该没有变化）
      const fieldNameMapping = new Map<string, string>();
      for (let i = 0; i < allFieldConfigs.length; i++) {
        const originalName = allFieldConfigs[i].name;
        const actualName = createdFields[i] ? createdFields[i].getName() : originalName;
        fieldNameMapping.set(originalName, actualName);
      }
      
      // 检查最终字段数量
      const finalFields = targetSheet.getFields();
      console.log('所有字段创建完成后的字段数量:', finalFields.length);
      console.log('最终字段列表:', finalFields.map((f: any) => f.getName()));
      
      // 合并所有记录
      for (const sheet of sheetsToMerge) {
        const records = await getAllRecordsFromSheet(sheet);
        
        // 转换记录数据以匹配新字段结构
        const transformedRecords = records.map(record => ({
          fields: transformRecordFieldsWithMapping(record, fieldMapping, fieldNameMapping, sheet.getName(), targetSheet)
        }));
        
        if (transformedRecords.length > 0) {
          await targetSheet.insertRecordsAsync(transformedRecords);
          totalRecords += transformedRecords.length;
        }
      }

      return {
        id: targetSheet.getId(),
        name: targetSheet.getName(),
        totalRecords,
        totalFields: finalFields.length,
        mergedSheets: sheetsToMerge.map(sheet => sheet.getName())
      };

    } else if (mergeMode === 'overwrite') {
      // 覆盖模式 - 智能字段映射
      if (!targetSheetId) {
        throw new Error('请选择目标数据表');
      }

      targetSheet = base.getSheet(targetSheetId);
      if (!targetSheet) {
        throw new Error('未找到目标数据表');
      }

      // 清空目标表的所有记录
      const existingRecords = await getAllRecordsFromSheet(targetSheet);
      if (existingRecords.length > 0) {
        const recordIds = existingRecords.map(record => record.getId());
        await targetSheet.deleteRecordsAsync(recordIds);
      }

      // 获取目标表原有字段
      const originalTargetFields = targetSheet.getFields();
      const originalFieldNames = originalTargetFields.map((field: any) => field.getName());
      
      // 使用独立的字段映射分析（覆盖模式）
      const fieldMapping = analyzeFieldMappingForOverwrite(sheetsToMerge);
      
      // 删除目标表的所有字段（除了主键字段）
      const remainingFields: any[] = [];
      for (const field of originalTargetFields) {
        try {
          // 检查是否是主键字段
          if (field.isPrimary && field.isPrimary()) {
            console.log('跳过主键字段:', field.getName());
            remainingFields.push(field);
          } else {
            targetSheet.deleteField(field.getId());
            console.log('成功删除字段:', field.getName());
          }
        } catch (error: any) {
          console.log('无法删除字段:', field.getName(), '错误:', error.message);
          remainingFields.push(field);
        }
      }
      
      // 获取剩余字段的名称，用于冲突检测
      const remainingFieldNames = new Set(remainingFields.map(f => f.getName()));
      console.log('剩余字段:', Array.from(remainingFieldNames));
      
      // 准备字段配置，包含来源字段
      const allFieldsToCreate = [
        ...fieldMapping.commonFields,  // 公共字段在前
        ...fieldMapping.nonCommonFields,  // 非公共字段在后
        {
          name: '合并来源',
          type: 'text'
        }
      ];
      
      for (const fieldConfig of allFieldsToCreate) {
        let fieldName = fieldConfig.name;
        let counter = 1;
        
        // 检查字段名冲突（包括剩余字段）
        while (remainingFieldNames.has(fieldName)) {
          fieldName = `${fieldConfig.name}_${counter}`;
          counter++;
        }
        
        try {
          console.log('创建字段:', fieldName, '类型:', fieldConfig.type);
          const field = targetSheet.insertField({
            name: fieldName,
            type: fieldConfig.type as any
          });
          
          // 如果是单选或多选字段，添加选项
          if ((fieldConfig.type === 'singleSelect' || fieldConfig.type === 'multiSelect') && fieldConfig.property?.choices) {
            const options = fieldConfig.property.choices.map((choice: any) => ({ name: choice.name }));
            if (options.length > 0) {
              console.log('为字段', fieldName, '添加选项:', options);
              field.addOptions(options);
            }
          }
        } catch (error: any) {
          console.error(`创建字段 "${fieldName}" 失败:`, error.message);
          // 如果字段创建失败，尝试使用备用名称
          const backupName = `${fieldConfig.name}_备用${counter}`;
          try {
            const field = targetSheet.insertField({
              name: backupName,
              type: fieldConfig.type as any
            });
            console.log('使用备用名称创建字段:', backupName);
          } catch (backupError: any) {
            console.error(`备用字段 "${backupName}" 也创建失败:`, backupError.message);
          }
        }
      }
      

      // 合并所有记录
      for (const sheet of sheetsToMerge) {
        const records = await getAllRecordsFromSheet(sheet);
        
        // 转换记录数据以匹配新的字段结构
        const transformedRecords = records.map(record => ({
          fields: transformRecordFieldsForOverwrite(record, fieldMapping, sheet.getName(), targetSheet)
        }));
        
        if (transformedRecords.length > 0) {
          await targetSheet.insertRecordsAsync(transformedRecords);
          totalRecords += transformedRecords.length;
        }
      }

      return {
        id: targetSheet.getId(),
        name: targetSheet.getName(),
        totalRecords,
        totalFields: allFieldsToCreate.length + 1, // +1 for source field
        mergedSheets: sheetsToMerge.map(sheet => sheet.getName())
      };

    } else if (mergeMode === 'append') {
      // 追加模式
      if (!targetSheetId) {
        throw new Error('请选择目标数据表');
      }

      targetSheet = base.getSheet(targetSheetId);
      if (!targetSheet) {
        throw new Error('未找到目标数据表');
      }

      // 添加来源字段（如果不存在）
      const fields = targetSheet.getFields();
      let sourceFieldExists = fields.some((field: any) => field.getName() === '合并来源');
      if (!sourceFieldExists) {
        targetSheet.insertField({
          name: '合并来源',
          type: 'text' as any
        });
      }

      // 追加所有记录
      for (const sheet of sheetsToMerge) {
        const records = await getAllRecordsFromSheet(sheet);
        
        // 转换记录数据以匹配目标字段结构
        const transformedRecords = records.map(record => ({
          fields: transformRecordFieldsForExistingSheet(record, targetSheet, sheet.getName())
        }));
        
        if (transformedRecords.length > 0) {
          await targetSheet.insertRecordsAsync(transformedRecords);
          addedRecords += transformedRecords.length;
        }
      }

      return {
        id: targetSheet.getId(),
        name: targetSheet.getName(),
        addedRecords,
        totalFields: targetSheet.getFields().length,
        mergedSheets: sheetsToMerge.map(sheet => sheet.getName())
      };
    }
    
  } catch (error: any) {
    throw new Error(`合并数据表失败: ${error.message}`);
  }
}

// 分析字段映射
function analyzeFieldMapping(sheets: any[]) {
  const fieldMap = new Map<string, { name: string; type: string; count: number; options?: any[] }>();
  
  console.log('开始分析字段映射，数据表数量:', sheets.length);
  
  // 统计所有字段
  sheets.forEach((sheet, sheetIndex) => {
    const sheetName = sheet.getName();
    const fields = sheet.getFields();
    console.log(`数据表 ${sheetIndex + 1} (${sheetName}) 有 ${fields.length} 个字段:`, fields.map((f: any) => f.getName()));
    
    fields.forEach((field: any) => {
      const fieldName = field.getName();
      const fieldType = field.getType();
      
      if (fieldMap.has(fieldName)) {
        const existing = fieldMap.get(fieldName)!;
        existing.count++;
        console.log(`字段 "${fieldName}" 已存在，计数增加到 ${existing.count}`);
        // 如果类型不同，优先选择更通用的类型
        if (existing.type !== fieldType) {
          existing.type = getMoreGeneralType(existing.type, fieldType);
        }
        
        // 如果是单选或多选字段，合并选项
        if (fieldType === 'singleSelect' || fieldType === 'multiSelect') {
          const fieldOptions = field.getOptions ? field.getOptions() : [];
          console.log(`字段 "${fieldName}" 现有选项:`, existing.options);
          console.log(`字段 "${fieldName}" 新选项:`, fieldOptions);
          existing.options = mergeFieldOptions(existing.options || [], fieldOptions);
          console.log(`字段 "${fieldName}" 合并后选项:`, existing.options);
        }
      } else {
        const fieldData: any = {
          name: fieldName,
          type: fieldType,
          count: 1,
          sheetNames: [sheetName] // 记录字段来源的表名
        };
        
        // 如果是单选或多选字段，保存选项
        if (fieldType === 'singleSelect' || fieldType === 'multiSelect') {
          fieldData.options = field.getOptions ? field.getOptions() : [];
        }
        
        fieldMap.set(fieldName, fieldData);
        console.log(`新字段 "${fieldName}" 添加到映射中`);
      }
    });
  });

  // 构建最终字段列表：保持第一个表的字段顺序，公共字段在前，独有字段在后
  const firstSheetFields = sheets[0].getFields();
  const firstSheetFieldNames = firstSheetFields.map((f: any) => f.getName());
  const finalFields: any[] = [];
  const processedFields = new Set<string>();
  
  // 1. 按第一个表的顺序添加公共字段
  firstSheetFieldNames.forEach(fieldName => {
    if (fieldMap.has(fieldName)) {
      const fieldData = fieldMap.get(fieldName)!;
      finalFields.push({
        name: fieldData.name,
        type: fieldData.type,
        property: fieldData.options ? { choices: fieldData.options } : undefined
      });
      processedFields.add(fieldName);
    }
  });
  
  // 2. 添加其他表中独有的字段（按表顺序）
  sheets.slice(1).forEach(sheet => {
    const fields = sheet.getFields();
    fields.forEach((field: any) => {
      const fieldName = field.getName();
      if (fieldMap.has(fieldName) && !processedFields.has(fieldName)) {
        const fieldData = fieldMap.get(fieldName)!;
        // 为独有字段重命名：字段名_表名
        const renamedFieldName = `${fieldName}_${sheet.getName()}`;
        finalFields.push({
          name: renamedFieldName,
          type: fieldData.type,
          property: fieldData.options ? { choices: fieldData.options } : undefined
        });
        processedFields.add(fieldName);
        console.log(`独有字段 "${fieldName}" 重命名为 "${renamedFieldName}"`);
      }
    });
  });

  return {
    fields: finalFields,
    fieldMap: fieldMap
  };
}

// 合并字段选项，去重并保持名称唯一
function mergeFieldOptions(existingOptions: any[], newOptions: any[]) {
  const optionMap = new Map<string, any>();
  
  // 添加现有选项
  existingOptions.forEach(option => {
    optionMap.set(option.name, option);
  });
  
  // 添加新选项，如果名称已存在则跳过
  newOptions.forEach(option => {
    if (!optionMap.has(option.name)) {
      optionMap.set(option.name, option);
    }
  });
  
  return Array.from(optionMap.values());
}

// 根据字段名称和选项名称查找匹配的选项ID
function findMatchingOptionId(fieldName: string, optionName: string, targetSheet: any) {
  try {
    const field = targetSheet.getField(fieldName);
    if (field && field.getOptions) {
      const options = field.getOptions();
      const matchingOption = options.find((option: any) => option.name === optionName);
      return matchingOption ? matchingOption.id : null;
    }
  } catch (error) {
    console.log(`查找字段 ${fieldName} 的选项时出错:`, error);
  }
  return null;
}

// 分析覆盖模式的字段映射
function analyzeFieldMappingForOverwrite(sheets: any[]) {
  const fieldMap = new Map<string, { name: string; type: string; count: number; options?: any[]; sheetNames: string[] }>();
  
  // 统计所有字段（只统计字段名，不考虑内容）
  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const fields = sheet.getFields();
    fields.forEach((field: any) => {
      const fieldName = field.getName();
      const fieldType = field.getType();
      
      if (fieldMap.has(fieldName)) {
        const existing = fieldMap.get(fieldName)!;
        existing.count++;
        if (!existing.sheetNames.includes(sheetName)) {
          existing.sheetNames.push(sheetName);
        }
        // 如果类型不同，优先选择更通用的类型
        if (existing.type !== fieldType) {
          existing.type = getMoreGeneralType(existing.type, fieldType);
        }
        
        // 如果是单选或多选字段，合并选项
        if (fieldType === 'singleSelect' || fieldType === 'multiSelect') {
          const fieldOptions = field.getOptions ? field.getOptions() : [];
          existing.options = mergeFieldOptions(existing.options || [], fieldOptions);
        }
      } else {
        const fieldData: any = {
          name: fieldName,
          type: fieldType,
          count: 1,
          sheetNames: [sheetName]
        };
        
        // 如果是单选或多选字段，保存选项
        if (fieldType === 'singleSelect' || fieldType === 'multiSelect') {
          fieldData.options = field.getOptions ? field.getOptions() : [];
        }
        
        fieldMap.set(fieldName, fieldData);
      }
    });
  });

  // 构建最终字段列表：保持第一个表的字段顺序，公共字段在前，独有字段在后
  const firstSheetFields = sheets[0].getFields();
  const firstSheetFieldNames = firstSheetFields.map((f: any) => f.getName());
  const commonFields: any[] = [];
  const nonCommonFields: any[] = [];
  const processedFields = new Set<string>();
  
  // 1. 按第一个表的顺序添加公共字段
  firstSheetFieldNames.forEach(fieldName => {
    if (fieldMap.has(fieldName)) {
      const fieldData = fieldMap.get(fieldName)!;
      const isCommon = fieldData.count > 1;
      
      if (isCommon) {
        commonFields.push({
          name: fieldName,
          type: fieldData.type,
          property: fieldData.options ? { choices: fieldData.options } : undefined
        });
      } else {
        // 非公共字段重命名：原字段名_sheet名
        const sheetName = fieldData.sheetNames[0];
        nonCommonFields.push({
          name: `${fieldName}_${sheetName}`,
          type: fieldData.type,
          property: fieldData.options ? { choices: fieldData.options } : undefined
        });
      }
      processedFields.add(fieldName);
    }
  });
  
  // 2. 添加其他表中独有的字段（按表顺序）
  sheets.slice(1).forEach(sheet => {
    const fields = sheet.getFields();
    fields.forEach((field: any) => {
      const fieldName = field.getName();
      if (fieldMap.has(fieldName) && !processedFields.has(fieldName)) {
        const fieldData = fieldMap.get(fieldName)!;
        // 为独有字段重命名
        const renamedFieldName = `${fieldName}_${sheet.getName()}`;
        nonCommonFields.push({
          name: renamedFieldName,
          type: fieldData.type,
          property: fieldData.options ? { choices: fieldData.options } : undefined
        });
        processedFields.add(fieldName);
      }
    });
  });

  return {
    commonFields,
    nonCommonFields,
    fieldMap
  };
}

// 转换覆盖模式的记录字段
function transformRecordFieldsForOverwrite(record: any, fieldMapping: any, sourceSheetName: string, targetSheet: any) {
  const transformedFields: Record<string, any> = {};
  
  // 添加来源字段
  transformedFields['合并来源'] = sourceSheetName;
  
  const originalFields = record.getCellValues();
  
  // 处理公共字段
  fieldMapping.commonFields.forEach((field: any) => {
    const fieldName = field.name;
    const originalValue = originalFields[fieldName];
    
    if (originalValue !== null && originalValue !== undefined) {
      // 处理单选和多选字段
      if (field.type === 'singleSelect' && originalValue && typeof originalValue === 'object') {
        const optionId = findMatchingOptionId(fieldName, originalValue.name, targetSheet);
        transformedFields[fieldName] = optionId;
      } else if (field.type === 'multiSelect' && Array.isArray(originalValue)) {
        const matchingIds = originalValue
          .filter(item => item && item.name)
          .map(item => findMatchingOptionId(fieldName, item.name, targetSheet))
          .filter(id => id !== null);
        transformedFields[fieldName] = matchingIds;
      } else {
        transformedFields[fieldName] = originalValue;
      }
    }
  });
  
  // 处理非公共字段（重命名的字段）
  fieldMapping.nonCommonFields.forEach((field: any) => {
    // 从重命名的字段名中提取原始字段名
    const originalFieldName = field.name.replace(/_[^_]+$/, ''); // 移除最后一个下划线后的部分
    const originalValue = originalFields[originalFieldName];
    
    if (originalValue !== null && originalValue !== undefined) {
      // 处理单选和多选字段
      if (field.type === 'singleSelect' && originalValue && typeof originalValue === 'object') {
        const optionId = findMatchingOptionId(field.name, originalValue.name, targetSheet);
        transformedFields[field.name] = optionId;
      } else if (field.type === 'multiSelect' && Array.isArray(originalValue)) {
        const matchingIds = originalValue
          .filter(item => item && item.name)
          .map(item => findMatchingOptionId(field.name, item.name, targetSheet))
          .filter(id => id !== null);
        transformedFields[field.name] = matchingIds;
      } else {
        transformedFields[field.name] = originalValue;
      }
    }
  });
  
  return transformedFields;
}

// 获取更通用的字段类型
function getMoreGeneralType(type1: string, type2: string): string {
  const typeHierarchy = {
    'text': 1,
    'number': 2,
    'date': 3,
    'singleSelect': 4,
    'multiSelect': 5,
    'attachment': 6,
    'link': 7,
    'currency': 8,
    'progress': 9,
    'rating': 10,
    'user': 11,
    'department': 12,
    'group': 13,
    'oneWayLink': 14,
    'twoWayLink': 15
  };
  
  const priority1 = typeHierarchy[type1 as keyof typeof typeHierarchy] || 999;
  const priority2 = typeHierarchy[type2 as keyof typeof typeHierarchy] || 999;
  
  // 返回优先级更高的类型（数字越小越通用）
  return priority1 <= priority2 ? type1 : type2;
}

// 获取数据表的所有记录
async function getAllRecordsFromSheet(sheet: any) {
  const allRecords: any[] = [];
  let hasMore = true;
  let cursor = '';
  
  while (hasMore) {
    const result = await sheet.getRecordsAsync({
      pageSize: 100,
      cursor: cursor
    });
    
    allRecords.push(...result.records);
    hasMore = result.hasMore;
    cursor = result.cursor;
  }
  
  return allRecords;
}

// 转换记录字段以匹配新字段结构
function transformRecordFields(record: any, fieldMapping: any, sourceSheetName: string, targetSheet: any) {
  const transformedFields: Record<string, any> = {};
  
  // 添加来源字段
  transformedFields['合并来源'] = sourceSheetName;
  
  // 转换其他字段
  const originalFields = record.getCellValues();
  
  fieldMapping.fields.forEach((field: any) => {
    const fieldName = field.name;
    const originalValue = originalFields[fieldName];
    
    if (originalValue !== null && originalValue !== undefined) {
      // 处理单选和多选字段
      if (field.type === 'singleSelect' && originalValue && typeof originalValue === 'object') {
        // 单选字段：根据名称匹配选项ID
        const optionId = findMatchingOptionId(fieldName, originalValue.name, targetSheet);
        transformedFields[fieldName] = optionId;
      } else if (field.type === 'multiSelect' && Array.isArray(originalValue)) {
        // 多选字段：根据名称数组匹配选项ID数组
        const matchingIds = originalValue
          .filter(item => item && item.name)
          .map(item => findMatchingOptionId(fieldName, item.name, targetSheet))
          .filter(id => id !== null);
        transformedFields[fieldName] = matchingIds;
      } else {
        // 其他字段类型直接赋值
        transformedFields[fieldName] = originalValue;
      }
    }
  });
  
  return transformedFields;
}

// 转换记录字段以匹配新字段结构（支持字段名映射）
function transformRecordFieldsWithMapping(record: any, fieldMapping: any, fieldNameMapping: Map<string, string>, sourceSheetName: string, targetSheet: any) {
  const transformedFields: Record<string, any> = {};
  
  // 添加来源字段
  transformedFields['合并来源'] = sourceSheetName;
  
  // 转换其他字段
  const originalFields = record.getCellValues();
  
  fieldMapping.fields.forEach((field: any) => {
    const originalFieldName = field.name;
    const actualFieldName = fieldNameMapping.get(originalFieldName) || originalFieldName;
    const originalValue = originalFields[originalFieldName];
    
    if (originalValue !== null && originalValue !== undefined) {
      // 处理单选和多选字段
      if (field.type === 'singleSelect' && originalValue && typeof originalValue === 'object') {
        // 单选字段：根据名称匹配选项ID
        const optionId = findMatchingOptionId(actualFieldName, originalValue.name, targetSheet);
        transformedFields[actualFieldName] = optionId;
      } else if (field.type === 'multiSelect' && Array.isArray(originalValue)) {
        // 多选字段：根据名称数组匹配选项ID数组
        const matchingIds = originalValue
          .filter(item => item && item.name)
          .map(item => findMatchingOptionId(actualFieldName, item.name, targetSheet))
          .filter(id => id !== null);
        transformedFields[actualFieldName] = matchingIds;
      } else {
        // 其他字段类型直接赋值
        transformedFields[actualFieldName] = originalValue;
      }
    }
  });
  
  return transformedFields;
}

// 转换记录字段以匹配现有数据表结构
function transformRecordFieldsForExistingSheet(record: any, targetSheet: any, sourceSheetName: string) {
  const transformedFields: Record<string, any> = {};
  
  // 添加来源字段
  transformedFields['合并来源'] = sourceSheetName;
  
  // 获取目标表的字段
  const targetFields = targetSheet.getFields();
  const originalFields = record.getCellValues();
  
  // 转换字段，只保留目标表中存在的字段
  targetFields.forEach((field: any) => {
    const fieldName = field.getName();
    const fieldType = field.getType();
    const originalValue = originalFields[fieldName];
    
    // 跳过合并来源字段，因为我们已经单独处理了
    if (fieldName === '合并来源') {
      return;
    }
    
    if (originalValue !== null && originalValue !== undefined) {
      // 处理单选和多选字段
      if (fieldType === 'singleSelect' && originalValue && typeof originalValue === 'object') {
        // 单选字段：根据名称匹配目标表中的选项ID
        const targetOptions = field.getOptions ? field.getOptions() : [];
        const matchingOption = targetOptions.find((option: any) => option.name === originalValue.name);
        transformedFields[fieldName] = matchingOption ? matchingOption.id : null;
      } else if (fieldType === 'multiSelect' && Array.isArray(originalValue)) {
        // 多选字段：根据名称数组匹配目标表中的选项ID数组
        const targetOptions = field.getOptions ? field.getOptions() : [];
        const matchingIds = originalValue
          .filter(item => item && item.name)
          .map(item => {
            const matchingOption = targetOptions.find((option: any) => option.name === item.name);
            return matchingOption ? matchingOption.id : null;
          })
          .filter(id => id !== null);
        transformedFields[fieldName] = matchingIds;
      } else {
        // 其他字段类型直接赋值
        transformedFields[fieldName] = originalValue;
      }
    }
  });
  
  return transformedFields;
}

// 2.注册指令「getActiveSheetName」
DingdocsScript.registerScript('getActiveSheetName', getActiveSheetName);

// 注册所有方法供UI层调用
DingdocsScript.registerScript('getActiveSheet', getActiveSheet);
DingdocsScript.registerScript('getAllSheets', getAllSheets);
DingdocsScript.registerScript('createSheet', createSheet);
DingdocsScript.registerScript('deleteSheet', deleteSheet);
DingdocsScript.registerScript('getSheetFields', getSheetFields);
DingdocsScript.registerScript('addField', addField);
DingdocsScript.registerScript('deleteField', deleteField);
DingdocsScript.registerScript('getRecords', getRecords);
DingdocsScript.registerScript('addRecord', addRecord);
DingdocsScript.registerScript('updateRecord', updateRecord);
DingdocsScript.registerScript('deleteRecord', deleteRecord);
DingdocsScript.registerScript('getDocumentInfo', getDocumentInfo);
DingdocsScript.registerScript('mergeSheets', mergeSheets);

export {};
