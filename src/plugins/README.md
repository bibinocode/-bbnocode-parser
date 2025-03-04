# Plugin 插件系统


- 系统会自动扫描 `src/plugins` 目录下的所有插件 后缀为 `.js` 或 `.mjs` 的文件
- 可以通过 `loadPlugins` 手动加载插件



## 插件编写结构

```js
// src/plugins/tableParser.js
export default {
  name: '表格解析插件',
  priority: 5,
  
  // 生命周期钩子
  onInit() {
    console.log('表格解析插件已初始化');
  },
  
  onBeforeProcess(data) {
    console.log('开始处理表格');
  },
  
  onAfterProcess(data) {
    console.log('表格处理完成');
  },
  
  onError(error, data) {
    console.error('表格处理出错', error);
  },
  
  onDestroy() {
    console.log('表格解析插件已卸载');
  },
  
  // 主处理逻辑
  async process(data) {
    // 表格解析逻辑
    data.content.tables = await extractTables(data.document);
    return data;
  },
  
  // 事件订阅
  subscribe: {
    'DOCUMENT_LOADED': (event) => {
      console.log('文档已加载，准备解析表格');
    }
  },
  
  // 元数据
  meta: {
    version: '1.0.0',
    author: 'bbnocode',
    description: '解析Word文档中的表格'
  }
};

// 辅助函数
async function extractTables(doc) {
  // 表格提取逻辑
  // ...
  return [];
}
```

## 事件系统

- 插件可以订阅和发布事件

```js
// 插件定义中订阅事件
{
  name: '我的插件',
  process: async (data) => { 

    // content 对象是解析后的存储对象 
    return data
  },
  subscribe: {
    'DOCUMENT_LOADED': (event) => {
      console.log('文档已加载');
    },
    'PARSE_COMPLETE': (event) => {
      console.log('解析完成，结果：', event.payload.result);
    }
  }
}

// 也可以在外部订阅事件
import { subscribeToEvents } from './src/index.js';

const unsubscribe = subscribeToEvents(['PARSE_START', 'PARSE_COMPLETE'], (event) => {
  console.log(`收到事件: ${event.type}`, event.payload);
});

// 不再需要监听时取消订阅
unsubscribe();
```

## 插件管理

- 支持动态启用/禁用和卸载插件：

```js
import { 
  registerPlugin, 
  unregisterPlugin, 
  setPluginEnabled,
  getPlugins 
} from './src/index.js';

// 注册插件
const name = registerPlugin(myPlugin);

// 禁用插件
setPluginEnabled(name, false);

// 启用插件
setPluginEnabled(name, true);

// 卸载插件
unregisterPlugin(name);

// 获取所有插件信息
const pluginList = getPlugins();
```

## 状态流监控

```js
import { getSystemState, getPluginRegistry, getEventBus } from './src/index.js';

// 监控系统状态
getSystemState().subscribe(state => {
  console.log('系统状态：', state);
});

// 监控插件注册表变化
getPluginRegistry().subscribe(plugins => {
  console.log('已加载的插件：', plugins);
});

// 监控所有事件
getEventBus().subscribe(event => {
  console.log('系统事件：', event);
});
```