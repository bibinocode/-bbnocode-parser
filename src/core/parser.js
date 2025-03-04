/**
 * @author bbnocode@outlook.com
 * @description 文档解析模块
 */

import { docx } from "docx4js";
import JSZip from "jszip";


import {
  EVENT_TYPES,
  getPlugins,
  loadPluginsFromDir,
  publishEvent
} from './plugin.js';


/**
 * 解析入口
 * @param {Buffer | string} docxData Word文档数据
 * @param {ParseOptions} options 解析选项
 * @returns {Promise<Object>} 解析结果
 * 
 * 
 * @typedef {Object} ParseOptions
 * @property {Boolean} autoLoadPlugins 是否自动加载插件
 * @property {String} pluginsDir 插件目录
 * @property {Boolean} stopOnError 是否在插件处理错误时停止解析
 */
async function parse(docxData, options = {}) {
  if (!docxData) {
    throw new Error('必须提供文档数据')
  }

  // 加载内置插件 plugins/**/*.js */
  await loadPluginsFromDir()

  // 如果没有插件,自动加载插件目录
  // const plugins = getPlugins()
  // if(plugins.length === 0 && options.autoLoadPlugins !== false){
  //   await loadPluginsFromDir(options?.pluginsDir)

  //   // 没有插件,则使用默认插件
  //   if(getPlugins().length === 0){
  //     registerPlugin(createDefaultPlugin())
  //   }
  // }

  publishEvent(EVENT_TYPES.PARSE_START, { options })

  // 解析文档

  try {

    /**
     * 这里考虑不用docx4js的方式,而是自己解析zip文件吧
     */
    const zip = await JSZip.loadAsync(docxData)
    const doc = await docx.load(docxData)

    publishEvent(EVENT_TYPES.PARSE_LOADED, { document: doc })

    let data = {
      document: doc,
      zip,
      content: {},
      ctx: {},
      options
    }


    // 按顺序执行所有启用的插件
    const enabledPlugins = getPlugins().filter(plugin => plugin.enabled)
    publishEvent(EVENT_TYPES.PLUGINS_PROCESSING, { count: enabledPlugins.length })

    for (const plugin of enabledPlugins) {

      publishEvent(EVENT_TYPES.PLUGIN_PROCESS_START, {
        plugin: plugin.name,
        priority: plugin.priority
      })

      try {
        // 安全调用钩子
        if (typeof plugin.onBeforeProcess === 'function') {
          await plugin.onBeforeProcess(data);
        }

        // process 是必须的
        data = await plugin.process(data);

        if (typeof plugin.onAfterProcess === 'function') {
          await plugin.onAfterProcess(data);
        }

        publishEvent(EVENT_TYPES.PLUGIN_PROCESS_COMPLETE, {
          plugin: plugin.name,
          success: true
        })
      } catch (error) {
        console.log("plugin", plugin);
        console.error(`插件 ${plugin.name} 执行出错:`, error);

        try {
          if (typeof plugin.onError === 'function') {
            await plugin.onError(error, data);
          }
        } catch (hookError) {
          console.error(`插件 ${plugin.name} 错误处理钩子执行失败:`, hookError);
        }

        publishEvent(EVENT_TYPES.PLUGIN_PROCESS_ERROR, {
          plugin: plugin.name,
          error: error.message
        });

        if (options.stopOnError) {
          publishEvent(EVENT_TYPES.PARSE_ERROR, {
            error: error.message,
            plugin: plugin.name
          });

          throw error;
        }
      }
    }

    publishEvent(EVENT_TYPES.PARSE_COMPLETE, {
      result: data.content
    })

    return data.content

  } catch (error) {
    publishEvent(EVENT_TYPES.PARSE_ERROR, {
      error: error.message
    })

    throw error
  }
}


export { parse };

