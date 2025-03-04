/**
 * @author bbnocode@outlook.com
 * @date 2025-03-04
 * @description 插件管理
 */
import fs from 'fs'
import path from 'path'
import { BehaviorSubject, filter, Subject, takeUntil } from "rxjs"
import { pathToFileURL } from 'url'


/**
 * @typedef {Object} EventTypes
 * @property {String} PLUGIN_REGISTERED 插件注册事件
 * @property {String} PLUGIN_UNREGISTERED 插件卸载事件
 * @property {String} PLUGIN_STATE_CHANGED 插件状态变化事件
 * @property {String} PLUGINS_LOAD_START 插件加载开始事件
 * @property {String} PLUGINS_LOAD_ERROR 插件加载错误事件
 * @property {String} PLUGINS_DISCOVERED 插件发现事件
 * @property {String} PLUGINS_LOAD_COMPLETE 插件加载完成事件
 * @property {String} PLUGINS_MANUAL_LOAD_COMPLETE 插件手动加载完成事件
 * @property {String} PLUGINS_PROCESSING 插件处理中事件
 * @property {String} PLUGIN_PROCESS_START 插件处理开始事件
 * @property {String} PLUGIN_PROCESS_COMPLETE 插件处理完成事件
 * @property {String} PLUGIN_PROCESS_ERROR 插件处理错误事件
 * @property {String} PARSE_START 解析开始事件
 * @property {String} PARSE_LOADED 文档解析完成
 * @property {String} PARSE_ERROR 解析错误事件
 * @property {String} PARSE_COMPLETE 解析完成事件
 */
const EVENT_TYPES = {
  /**
   * 插件注册事件
   */
  PLUGIN_REGISTERED: 'PLUGIN_REGISTERED',
  /**
   * 插件卸载事件
   */
  PLUGIN_UNREGISTERED: 'PLUGIN_UNREGISTERED',
  /**
   * 插件状态变化事件
   */
  PLUGIN_STATE_CHANGED: 'PLUGIN_STATE_CHANGED',
  /**
   * 插件加载开始事件
   */
  PLUGINS_LOAD_START: 'PLUGINS_LOAD_START',
  /**
   * 插件加载错误事件
   */
  PLUGINS_LOAD_ERROR: 'PLUGINS_LOAD_ERROR',
  /**
   * 插件发现事件
   */
  PLUGINS_DISCOVERED: 'PLUGINS_DISCOVERED',
  /**
   * 插件加载完成事件
   */
  PLUGINS_LOAD_COMPLETE: 'PLUGINS_LOAD_COMPLETE',
  /**
   * 插件手动加载完成事件
   */
  PLUGINS_MANUAL_LOAD_COMPLETE: 'PLUGINS_MANUAL_LOAD_COMPLETE',
  /**
   * 插件处理中事件
   */
  PLUGINS_PROCESSING: 'PLUGINS_PROCESSING',
  /**
   * 插件处理开始事件
   */
  PLUGIN_PROCESS_START: 'PLUGIN_PROCESS_START',
  /**
   * 插件处理完成事件
   */
  PLUGIN_PROCESS_COMPLETE: 'PLUGIN_PROCESS_COMPLETE',
  /**
   * 插件处理错误事件
   */
  PLUGIN_PROCESS_ERROR: 'PLUGIN_PROCESS_ERROR',
  /**
   * 解析开始事件
   */
  PARSE_START: 'PARSE_START',
  /**
   * 文档解析完成
   */
  PARSE_LOADED: 'PARSE_LOADED',
  /**
   * 解析错误事件
   */
  PARSE_ERROR: 'PARSE_ERROR',
  /**
   * 解析完成事件
   */
  PARSE_COMPLETE: 'PARSE_COMPLETE',
}

/**
 * 插件存储
 */
const Plugins= []

/**
 * 事件总线
 */
const eventBus = new Subject()


/**
 * 插件状态管理
 */
const pluginRegistry = new BehaviorSubject()

/**
 * 系统状态
 */
const systemState = new BehaviorSubject({
  isProcessing:false,
  currentPhase:null
})



/**
 * 发布事件
 * @param {String} eventType 事件类型
 * @param {Object} payload 载体
 */
function publishEvent(eventType,payload = {}){
  eventBus.next({
    type:eventType,
    timestamp:Date.now(),
    payload
  })
}


/**
 * 订阅指定类型事件
 * @param {String|Array} eventType 事件类型
 * @param {Function} handler 处理器
 * @returns {Function} 取消订阅函数
 */
function subscribeToEvents(eventType,handler){
  const type = Array.isArray(eventType) ? eventType : [eventType]
  const unsubscribe = new Subject()

  eventBus.pipe(
    filter(event => type.includes(event.type)),
    takeUntil(unsubscribe)
  ).subscribe(handler)

  return ()=> unsubscribe.next()
}


/**
 * 注册插件事件
 * @param {Plugin} plugin 插件对象
 * @param {String} 插件名称
 * 
 * @typedef {Object} Plugin
 * @property {String} name 插件名称
 * @property {Object} meta 元数据
 * @property {Number} priority 优先级
 * @property {Function} process 处理函数
 * @property {Boolean} enabled 是否启用
 * @property {Function} onBeforeProcess 处理前回调
 * @property {Function} onAfterProcess 处理后回调
 * @property {Function} onError 错误处理
 * @property {Function} onDestroy 销毁函数
 * 
 * 
 */
function registerPlugin(plugin){
  if(!plugin || typeof plugin.process !== 'function'){
    throw new Error('请提供process处理函数')
  }

  const pluginName = plugin.name || `plugin_${Plugins.length}`

  const fullPlugin = {
    id: `plugin_${Date.now()}_${Math.floor(Math.random() * 1000)}`,
    name:pluginName,
    priority: plugin.priority || 10,
    process: plugin.process,
    enabled:plugin.enabled || true,
    onInit:plugin.onInit || (()=>{}),
    onBeforeProcess:plugin.onBeforeProcess || (()=>{}),
    onAfterProcess:plugin.onAfterProcess || (()=>{}),
    onError:plugin.onError || (()=>{}),
    onDestroy:plugin.onDestroy || (()=>{}),
    subscriptions:[], // 插件可以订阅的事件
    meta:plugin.meta || {}
  }

  try {
    fullPlugin.onInit()
  } catch (error) {
    console.error(`插件 ${fullPlugin.name} 初始化失败:`, error);
  }

  // 插件订阅
  if (plugin.subscribe && typeof plugin.subscribe === 'object') {
    for (const [eventType, handler] of Object.entries(plugin.subscribe)) {
      if (typeof handler === 'function') {
        const unsubscribe = subscribeToEvents(eventType, handler);
        fullPlugin.subscriptions.push(unsubscribe);
      }
    }
  }

  Plugins.push(fullPlugin)
  // 根据优先级排序
  Plugins.sort((a, b) => a.priority - b.priority);
  // 更新注册表
  pluginRegistry.next([...Plugins])

  // 发布注册事件
  publishEvent(EVENT_TYPES.PLUGIN_REGISTERED,{
    plugin:fullPlugin
  })

  return pluginName
}


/**
 * 卸载插件
 * @param {String} pluginNameOrId 插件名称或ID
 */
function unregisterPlugin(pluginNameOrId){
  const index = Plugins.findIndex(p => p.name === pluginNameOrId || p.id === pluginNameOrId);

  if(index === -1){
    console.warn(`未找到插件: ${pluginNameOrId}`);
    return false;
  }

  const plugin = Plugins[index]

  try {
    plugin.onDestroy()
  } catch (error) {
    console.error(`插件 ${plugin.name} 销毁时出错:`, error);
  }

  // 取消所有订阅
  for (const unsubscribe of plugin.subscriptions) {
    unsubscribe();
  }

  Plugins.splice(index,1)
  pluginRegistry.next([...Plugins])

  publishEvent(EVENT_TYPES.PLUGIN_UNREGISTERED,{
    plugin:plugin
  })

  return true
}


/**
 * 启用或禁用插件
 * @param {String} pluginNameOrId 插件名称或ID
 * @param {Boolean} enabled 是否启用
 * @returns {Boolean} 操作是否成功
 */
function setPluginEnabled(pluginNameOrId, enabled = true){
  const plugin = Plugins.findIndex(p => p.name === pluginNameOrId || p.id === pluginNameOrId);

  if(!plugin){
    console.warn(`未找到插件: ${pluginNameOrId}`);
    return false;
  }

  plugin.enabled = !!enabled;
  pluginRegistry.next([...Plugins])

  publishEvent(EVENT_TYPES.PLUGIN_STATE_CHANGED,{
    plugin:plugin,
    enabled:plugin.enabled
  })

  return true
}


/**
 * 扫描并加载插件
 * @param {String} pluginDir 插件目录 默认为 src/plugins
 * @returns {Promise<Array>} 加载的插件名称列表
 */
async function loadPluginsFromDir(pluginDir = 'src/plugins'){
  publishEvent(EVENT_TYPES.PLUGINS_LOAD_START, { directory: pluginDir });

  if(pluginDir === 'src/plugins'){
    pluginDir = path.resolve(process.cwd(), pluginDir)
  }
  if (!fs.existsSync(pluginDir)) {
    console.warn(`插件目录 ${pluginDir} 不存在`);
    publishEvent(EVENT_TYPES.PLUGINS_LOAD_ERROR, { 
      directory: pluginDir,
      error: '目录不存在'
    });
    return [];
  }

  // 使用glob 递归查找

  const {glob} = await import('glob')
  const pluginFiles = await glob("*/index.{js,mjs}", {cwd: pluginDir,dot:false})
  
  const loadPlugins = []

  publishEvent(EVENT_TYPES.PLUGINS_DISCOVERED, { 
    directory: pluginDir,
    files: pluginFiles
  });


  for (const file of pluginFiles){
    try {
      const filePath = path.resolve(pluginDir,file)
      const fileUrl = pathToFileURL(filePath).href
      const plugin = await import(fileUrl)

      // 支持默认导出或命名导出
      const pluginModule = plugin.default || plugin

      if(typeof pluginModule === 'function'){
        // 如果导出的是函数,需要获取插件实例
        const pluginInstance = pluginModule()
        const pluginName = registerPlugin(pluginInstance)
        loadPlugins.push(pluginName)
      }else if(typeof pluginModule === 'object' && typeof pluginModule.process === 'function' ){
        // 如果导出的是对象,且有process函数,则直接注册
        const pluginName = registerPlugin(pluginModule);
        loadPlugins.push(pluginName);
      }else{
        console.warn(`加载插件 ${file} 格式不正确,跳过`)
        publishEvent(EVENT_TYPES.PLUGINS_LOAD_ERROR, { 
          file: file,
          error: '格式不正确'
        });
      }

    } catch (error) {
      console.error(`加载插件 ${file} 失败:`, error);
      publishEvent(EVENT_TYPES.PLUGINS_LOAD_ERROR, { 
        file: file,
        error: error.message
      });
    }
  }

  publishEvent(EVENT_TYPES.PLUGINS_LOAD_COMPLETE, { 
    directory: pluginDir,
    count: loadPlugins.length
  });
  console.log(`已从 ${pluginDir} 加载 ${loadPlugins.length} 个插件`);
  return loadPlugins
}


/**
 * 手动加载插件
 * @param {String|Array} pluginPaths 插件文件路径或路径数组
 * @returns {Promise<Array>} 加载的插件名称列表
 */
async function loadPlugins(pluginPaths) {
  if (!pluginPaths) return [];
  
  const paths = Array.isArray(pluginPaths) ? pluginPaths : [pluginPaths];
  const loadedPlugins = [];

  publishEvent(EVENT_TYPES.PLUGINS_LOAD_START, { 
    directory: pluginPaths,
    count: paths.length
  });
  
  for (const pluginPath of paths) {
    try {
      const fileUrl = pathToFileURL(path.resolve(pluginPath)).href;
      const plugin = await import(fileUrl);
      // 支持默认导出或命名导出
      const pluginModule = plugin.default || plugin;
      
      if (typeof pluginModule === 'function') {
        // 如果导出的是函数，执行它获取插件实例
        const pluginInstance = pluginModule();
        const pluginName = registerPlugin(pluginInstance);
        loadedPlugins.push(pluginName);
      } else if (typeof pluginModule === 'object' && typeof pluginModule.process === 'function') {
        // 如果导出的是插件对象，直接注册
        const pluginName = registerPlugin(pluginModule);
        loadedPlugins.push(pluginName);
      } else {
        console.warn(`插件文件 ${pluginPath} 格式不正确，已跳过`);
        publishEvent(EVENT_TYPES.PLUGINS_LOAD_ERROR, { 
          file: pluginPath,
          error: '格式不正确'
        });
      }
    } catch (error) {
      console.error(`加载插件 ${pluginPath} 失败:`, error);
      publishEvent(EVENT_TYPES.PLUGINS_LOAD_ERROR, { 
        file: pluginPath,
        error: error.message
      });
    }
  }
  
  publishEvent(EVENT_TYPES.PLUGINS_MANUAL_LOAD_COMPLETE, { 
    paths,
    loaded: loadedPlugins
  });
  
  return loadedPlugins;
}


/**
 * 创建默认插件
 * @returns {Object} 默认插件
 */
function createDefaultPlugin() {
  return {
    name: 'default',
    priority: 100,
    process: async (data) => {
      // 默认解析逻辑
      const doc = data.document;
      data.content.text = await doc.getText();
      return data;
    },
    onInit: () => {
      console.log('默认插件已初始化');
    },
    subscribe: {
      PARSE_START: (event) => {
        console.log('文档解析开始');
      },
      PARSE_COMPLETE: (event) => {
        console.log('文档解析完成');
      }
    }
  };
}


/**
 * 获取所有已加载的插件信息
 * @returns {Array} 插件信息列表
 */
function getPlugins() {
  return Plugins
}


/**
 * 获取系统状态的Observable
 * @returns {Observable} 系统状态Observable
 */
function getSystemState() {
  return systemState.asObservable();
}

/**
 * 获取插件注册表的Observable
 * @returns {Observable} 插件注册表Observable
 */
function getPluginRegistry() {
  return pluginRegistry.asObservable();
}

/**
 * 获取事件总线的Observable
 * @returns {Observable} 事件总线Observable
 */
function getEventBus() {
  return eventBus.asObservable();
}

// 导出模块
export {
  createDefaultPlugin, EVENT_TYPES, getEventBus, getPluginRegistry, getPlugins,
  getSystemState, loadPlugins, loadPluginsFromDir, publishEvent, registerPlugin, setPluginEnabled, subscribeToEvents, unregisterPlugin
}

