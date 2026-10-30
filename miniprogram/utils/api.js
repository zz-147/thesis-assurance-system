/**
 * tas 文档接龙 · 后端接口封装（原生 wx.request / wx.uploadFile，无第三方依赖）。
 *
 * 后端（tas 项目根目录）启动：
 *   python -m tas web --port 8765
 *
 * BASE_URL 说明：
 *   1) 微信开发者工具模拟器：保持 http://127.0.0.1:8765 即可（本机后端）。
 *   2) 真机预览：改成电脑在局域网里的 IPv4，例如 http://192.168.1.8:8765；
 *      手机和电脑连同一个 Wi-Fi，后端用 --host 0.0.0.0 启动，Windows 防火墙放行 8765。
 *   3) 上线：必须改成 https:// 的已备案域名，并在小程序后台
 *      「开发管理 → 开发设置 → 服务器域名」把同一个域名加到 request 与 uploadFile 合法域名。
 *      小程序不允许请求 http:// 或裸 IP（开发工具里勾选“不校验合法域名”后本地调试可用）。
 *
 * 错误约定：后端 HTTP 4xx/5xx 返回 { "ok": false, "error": "可读中文" }；
 * 本文件会把 error 字段抽成 Error.message，页面直接用中文提示即可。
 */

// 后端地址（末尾不要带 /）
// 真机调试必须用电脑局域网 IP（127.0.0.1 在手机上指向手机自己）。
// 可在小程序首页「网络自检」里点「修改地址」直接改，保存在本机，无需改代码。
var DEFAULT_BASE_URL = 'http://10.127.173.183:8765';

function normalizeBaseUrl(url) {
  return String(url || '').trim().replace(/\/+$/, '');
}

var BASE_URL = DEFAULT_BASE_URL;
try {
  var savedBaseUrl = wx.getStorageSync('tas_base_url');
  if (savedBaseUrl) {
    BASE_URL = normalizeBaseUrl(savedBaseUrl) || DEFAULT_BASE_URL;
  }
} catch (err) {
  // 读取缓存失败时用默认地址
}

/** 运行时修改后端地址（存本机）；返回规范化后的地址，非法地址抛错。 */
function setBaseUrl(url) {
  var next = normalizeBaseUrl(url);
  if (!/^https?:\/\/.+/.test(next)) {
    throw new Error('地址要以 http:// 或 https:// 开头，例如 http://192.168.1.8:8765');
  }
  BASE_URL = next;
  try {
    wx.setStorageSync('tas_base_url', next);
  } catch (err) {
    // 存不进去也不影响本次使用
  }
  return next;
}

// 小程序 wx.uploadFile 单文件上限为 10MB，页面在选择文件后按这个值校验
var MAX_UPLOAD_BYTES = 10 * 1024 * 1024;

// 普通请求超时；审核要逐份检查文档，单独放宽
var DEFAULT_TIMEOUT = 30000;
var REVIEW_TIMEOUT = 180000;
var REPORT_TIMEOUT = 60000;

/** 拼接完整 URL（容忍 BASE_URL 末尾误写的 /）。 */
function apiUrl(path) {
  var base = String(BASE_URL || '').replace(/\/+$/, '');
  return base + path;
}

/** wx.request / wx.uploadFile 的 data 可能是对象或字符串，统一尝试解析成对象。 */
function parseBody(raw) {
  if (typeof raw !== 'string') {
    return raw;
  }
  if (!raw) {
    return null;
  }
  try {
    return JSON.parse(raw);
  } catch (err) {
    return raw;
  }
}

/** 从响应体里提取后端的中文错误信息。 */
function pickErrorMessage(data) {
  if (data && typeof data === 'object') {
    return data.error || data.message || '';
  }
  if (typeof data === 'string' && data) {
    try {
      var parsed = JSON.parse(data);
      return (parsed && (parsed.error || parsed.message)) || '';
    } catch (err) {
      return '';
    }
  }
  return '';
}

/** HTTP 非 2xx 时的中文错误。 */
function httpErrorMessage(statusCode, data) {
  var message = pickErrorMessage(data);
  if (message) {
    return message;
  }
  if (statusCode === 404) {
    return '接口不存在（HTTP 404）：请确认 tas 后端已启动，且版本包含文档接龙接口（/api/collections）';
  }
  if (statusCode === 413) {
    return '文件太大（HTTP 413）：小程序单文件上限 10MB，请压缩后重试';
  }
  if (statusCode >= 500) {
    return '后端处理失败（HTTP ' + statusCode + '），请查看后端终端日志';
  }
  return '请求失败（HTTP ' + statusCode + '）';
}

/** 网络层失败（后端没启动 / IP 不对 / 超时）时的中文错误。 */
function networkErrorMessage(err) {
  var raw = (err && err.errMsg) || '';
  if (raw.indexOf('timeout') >= 0) {
    return '请求超时：文档较多或网络较慢，请稍后重试';
  }
  return '网络连接失败：请确认后端已启动（' + BASE_URL + '）；真机需与电脑在同一局域网，并把 utils/api.js 的 BASE_URL 改成电脑 IP，详见 miniprogram/README.md';
}

/** 通用 JSON 请求，resolve 后端 JSON；失败 reject 一个带中文 message 的 Error。 */
function request(path, options) {
  var options_ = options || {};
  return new Promise(function (resolve, reject) {
    wx.request({
      url: apiUrl(path),
      method: options_.method || 'GET',
      data: options_.data,
      timeout: options_.timeout || DEFAULT_TIMEOUT,
      header: {
        'content-type': 'application/json'
      },
      success: function (res) {
        var data = parseBody(res.data);
        if (res.statusCode >= 200 && res.statusCode < 300) {
          if (data && data.ok === false) {
            reject(new Error(data.error || '请求失败'));
            return;
          }
          resolve(data);
          return;
        }
        reject(new Error(httpErrorMessage(res.statusCode, data)));
      },
      fail: function (err) {
        reject(new Error(networkErrorMessage(err)));
      }
    });
  });
}

/**
 * 上传文件（multipart/form-data），resolve 后端 JSON。
 * fields 会作为 multipart 的普通字段提交（提交文档时传 { student: "张三" }）。
 */
function uploadFile(path, filePath, fields) {
  var formData = fields || {};
  return new Promise(function (resolve, reject) {
    wx.uploadFile({
      url: apiUrl(path),
      filePath: filePath,
      name: 'file',
      formData: formData,
      success: function (res) {
        var data = parseBody(res.data);
        if (res.statusCode >= 200 && res.statusCode < 300) {
          if (data && data.ok === false) {
            reject(new Error(data.error || '上传失败'));
            return;
          }
          resolve(data);
          return;
        }
        reject(new Error(httpErrorMessage(res.statusCode, data)));
      },
      fail: function (err) {
        reject(new Error(networkErrorMessage(err)));
      }
    });
  });
}

var api = {
  MAX_UPLOAD_BYTES: MAX_UPLOAD_BYTES,

  // 当前后端地址（可在「网络自检」里通过 setBaseUrl 运行时修改，存本机）
  getBaseUrl: function () {
    return BASE_URL;
  },
  setBaseUrl: setBaseUrl,
  resetBaseUrl: function () {
    try {
      wx.removeStorageSync('tas_base_url');
    } catch (err) {
      // ignore
    }
    BASE_URL = DEFAULT_BASE_URL;
    return BASE_URL;
  },

  // GET /api/profiles → { default, profiles: [{ id, name, discipline }] }
  getProfiles: function () {
    return request('/api/profiles');
  },

  // GET /api/diag → { ok, client, server_host, server_port, base_url, verbose }
  // 真机网络排查：返回服务器看到的手机 IP；能通说明网络层没问题。
  diag: function () {
    return request('/api/diag', { timeout: 8000 });
  },

  // GET /api/collections → { ok, collections: [{ code, title, owner, ruleset, note, due, created_at, submission_count }] }
  listCollections: function () {
    return request('/api/collections');
  },

  // POST /api/collections (JSON) → { ok, collection: { code, ... } }
  // payload: { title, owner, ruleset, note, due }
  createCollection: function (payload) {
    return request('/api/collections', {
      method: 'POST',
      data: payload
    });
  },

  // GET /api/collections/{code} → { ok, collection, submissions: [{ id, student, filename, created_at, counts, status }] }
  getCollection: function (code) {
    return request('/api/collections/' + encodeURIComponent(code));
  },

  // POST /api/collections/{code}/submit (multipart: student + file) → { ok, submission }
  submitDocument: function (code, student, filePath) {
    return uploadFile('/api/collections/' + encodeURIComponent(code) + '/submit', filePath, {
      student: student
    });
  },

  // POST /api/collections/{code}/review → { ok, summary: { total, checked, failed, pass_rate, counts, top_rules }, submissions }
  // 审核不需要请求体，后端按 code 重新检查全部文档
  reviewCollection: function (code) {
    return request('/api/collections/' + encodeURIComponent(code) + '/review', {
      method: 'POST',
      timeout: REVIEW_TIMEOUT
    });
  },

  // GET /api/collections/{code}/submissions/{sid}/report → Report.to_dict 完整报告 JSON
  getSubmissionReport: function (code, submissionId) {
    return request(
      '/api/collections/' + encodeURIComponent(code) + '/submissions/' + encodeURIComponent(submissionId) + '/report',
      { timeout: REPORT_TIMEOUT }
    );
  }
};

module.exports = api;
