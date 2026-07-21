/**
 * 长舟编程学习平台 - CDN配置
 *
 * 图片存储在腾讯云COS，通过此地址加载。
 * 格式: https://<bucket>.cos.<region>.myqcloud.com/
 *
 * 修改此地址即可切换图片源，无需改动其他文件。
 * 留空则回退到本地相对路径。
 */
var COS_CDN_BASE = 'https://changzhou-ppt-1456685753.cos.ap-guangzhou.myqcloud.com/';
