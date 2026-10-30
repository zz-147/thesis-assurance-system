var api = require('../../utils/api');

// 规则编号 → 中文名（与 tas 仓库 rules/rules.json 的 name 字段一致，仅用于展示；
// 后端 /api/collections/{code}/review 的 top_rules 只给规则编号 + 次数）
var RULE_NAMES = {
  'PAGE-001': '纸张为 A4',
  'PAGE-002': '页边距、装订线、页眉页脚距边界',
  'PAGE-003': '文档网格',
  'PAGE-004': '分节与页眉页脚继承',
  'BODY-001': '正文字体',
  'BODY-002': '正文行距',
  'BODY-003': '正文首行缩进',
  'BODY-004': '正文两端对齐',
  'HEAD-001': '一级标题',
  'HEAD-002': '二级标题',
  'HEAD-003': '三级标题',
  'HEAD-004': '四级及以下标题',
  'HEAD-005': '文经管标题体系',
  'HEAD-006': '标题编号连续',
  'HEAD-007': '前言/结论/致谢标题',
  'TITLE-001': '中文题目',
  'TITLE-002': '英文题目',
  'ABS-001': '中文摘要标题',
  'ABS-002': '中文摘要正文',
  'ABS-003': '英文摘要标题',
  'ABS-004': '英文摘要正文',
  'KW-001': '中文关键词',
  'KW-002': '英文关键词',
  'TOC-001': '目录标题与条目样式',
  'TOC-002': '目录完整性',
  'HF-001': '页眉样式',
  'HF-002': '页眉文字',
  'HF-003': '页码格式',
  'FIG-001': '图题',
  'FIG-002': '图编号连续性',
  'TAB-001': '三线表样式',
  'TAB-002': '表题与表内文字',
  'EQ-001': '公式与公式编号',
  'REF-001': '参考文献标题',
  'REF-002': '参考文献条目格式',
  'REF-003': '正文引用标注',
  'STRUCT-001': '装订顺序完整',
  'MISC-001': '数字用法'
};

var SEVERITY_LABELS = {
  error: '错误',
  warning: '警告',
  info: '提示'
};

// 报告里最多展示多少条问题明细（完整报告见后端 JSON）
var MAX_SHOW_ISSUES = 30;

function pad2(value) {
  return value < 10 ? '0' + value : String(value);
}

function todayString() {
  var now = new Date();
  return now.getFullYear() + '-' + pad2(now.getMonth() + 1) + '-' + pad2(now.getDate());
}

function nowText() {
  var now = new Date();
  return now.getFullYear() + '-' + pad2(now.getMonth() + 1) + '-' + pad2(now.getDate()) +
    ' ' + pad2(now.getHours()) + ':' + pad2(now.getMinutes());
}

/** "2026-10-26T21:55:00+08:00" → "2026-10-26 21:55"；解析不了就原样返回。 */
function formatTime(value) {
  if (!value) {
    return '';
  }
  var text = String(value).replace('T', ' ');
  text = text.replace(/\.\d+/, '');
  text = text.replace(/([+-]\d{2}:?\d{2}|Z)$/, '');
  return text.length >= 16 ? text.slice(0, 16) : text;
}

/** Windows 路径 / URL 都取最后一段文件名。 */
function baseName(file) {
  var text = String(file || '');
  var slash = text.lastIndexOf('/');
  var backslash = text.lastIndexOf('\\');
  var index = Math.max(slash, backslash);
  return index >= 0 ? text.slice(index + 1) : text;
}

/** 0.66 → "66%"；也兼容后端直接给 66 这种百分比。 */
function formatPassRate(value) {
  var num = Number(value);
  if (!isFinite(num)) {
    return '—';
  }
  var percent = num <= 1 ? num * 100 : num;
  percent = Math.round(percent);
  if (percent < 0) {
    percent = 0;
  }
  if (percent > 100) {
    percent = 100;
  }
  return percent + '%';
}

function formatSize(bytes) {
  var num = Number(bytes || 0);
  if (num >= 1048576) {
    return (num / 1048576).toFixed(1) + 'MB';
  }
  if (num >= 1024) {
    return Math.round(num / 1024) + 'KB';
  }
  return num + 'B';
}

function pickCounts(counts) {
  var source = counts || {};
  return {
    error: Number(source.error || 0),
    warning: Number(source.warning || 0),
    info: Number(source.info || 0)
  };
}

function countsText(counts) {
  return '错误 ' + counts.error + ' · 警告 ' + counts.warning + ' · 提示 ' + counts.info;
}

function statusFromCounts(counts) {
  if (counts.error > 0) {
    return '未通过';
  }
  if (counts.warning > 0) {
    return '有警告';
  }
  return '通过';
}

function statusClass(status) {
  var text = String(status || '');
  if (text.indexOf('未通过') >= 0 || text.indexOf('失败') >= 0) {
    return 'tag-fail';
  }
  if (text.indexOf('通过') >= 0) {
    return 'tag-pass';
  }
  return 'tag-warn';
}

/** 接口提交记录 → 页面展示字段。 */
function formatSubmission(item) {
  var hasCounts = !!(item.counts && typeof item.counts === 'object');
  var counts = pickCounts(item.counts);
  var status = item.status || statusFromCounts(counts);
  return {
    id: item.id || '',
    student: item.student || '匿名',
    filename: item.filename || '未命名.docx',
    createdAt: formatTime(item.created_at),
    counts: counts,
    countsText: hasCounts ? countsText(counts) : '未能检查（文件可能损坏，可让同学重新提交）',
    status: status,
    statusClass: statusClass(status),
    error: counts.error,
    warning: counts.warning,
    info: counts.info
  };
}

/** 接口 collection → 详情页展示字段。 */
function formatCollectionDetail(item, rulesetNames) {
  var count = Number(item.submission_count || 0);
  var due = item.due ? String(item.due) : '';
  var dueDate = due.length >= 10 ? due.slice(0, 10) : due;
  var ruleset = item.ruleset || '';
  return {
    code: item.code || '',
    title: item.title || '(未命名接龙)',
    owner: item.owner || '匿名',
    ruleset: ruleset,
    rulesetName: rulesetNames[ruleset] || ruleset || '默认规范',
    note: item.note || '',
    due: due,
    dueText: dueDate ? '截止 ' + dueDate : '不限截止',
    duePast: !!dueDate && dueDate < todayString(),
    createdAt: formatTime(item.created_at),
    submissionCount: count,
    submissionText: count + ' 人已提交'
  };
}

/** 把 /review 返回的 top_rules（[["BODY-004", 12], ...]）转成可读列表。 */
function buildTopRules(rows) {
  var list = [];
  if (!Array.isArray(rows)) {
    return list;
  }
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var rule = '';
    var count = 0;
    if (Array.isArray(row)) {
      rule = String(row[0] || '');
      count = Number(row[1] || 0);
    } else if (row && typeof row === 'object') {
      rule = String(row.rule || row.rule_id || '');
      count = Number(row.count || 0);
    }
    if (!rule) {
      continue;
    }
    list.push({
      rule: rule,
      count: count,
      label: rule + (RULE_NAMES[rule] ? ' ' + RULE_NAMES[rule] : '')
    });
  }
  return list;
}

/** Report.to_dict 完整报告 → 页面展示用的摘要。 */
function normalizeReport(payload, submissionId) {
  var report = payload && payload.report ? payload.report : (payload || {});
  var counts = pickCounts(report.counts);
  var stats = report.stats || {};
  var rules = report.rules || {};
  // 后端可能在报告里附带 submission（原始文件名），有就优先展示
  var reportSubmission = report.submission && typeof report.submission === 'object' ? report.submission : {};

  var issues = [];
  var rawIssues = Array.isArray(report.issues) ? report.issues : [];
  var limit = Math.min(rawIssues.length, MAX_SHOW_ISSUES);
  for (var i = 0; i < limit; i++) {
    var issue = rawIssues[i] || {};
    var severity = issue.severity || 'info';
    issues.push({
      key: i,
      level: SEVERITY_LABELS[severity] || '提示',
      levelClass: severity === 'error' ? 'level-error' : (severity === 'warning' ? 'level-warning' : 'level-info'),
      rule: issue.rule_id || '',
      title: issue.title || '',
      locator: issue.locator || '',
      message: issue.message || ''
    });
  }

  var byRule = [];
  var issuesByRule = report.issues_by_rule || {};
  for (var ruleId in issuesByRule) {
    if (Object.prototype.hasOwnProperty.call(issuesByRule, ruleId)) {
      byRule.push({
        rule: ruleId,
        count: Number(issuesByRule[ruleId] || 0),
        label: ruleId + (RULE_NAMES[ruleId] ? ' ' + RULE_NAMES[ruleId] : '')
      });
    }
  }
  byRule.sort(function (a, b) {
    return b.count - a.count;
  });

  return {
    submissionId: submissionId,
    fileText: reportSubmission.filename || baseName(report.file || ''),
    profileLabel: report.profile_label || report.profile || '默认规范',
    generatedAt: formatTime(report.generated_at),
    counts: counts,
    error: counts.error,
    warning: counts.warning,
    info: counts.info,
    countsText: countsText(counts),
    rulesText: '检查规则：本次 ' + Number(rules.checked || 0) + ' / 共 ' + Number(rules.total || 0) + ' 条',
    statsText: '文档：段落 ' + Number(stats.paragraphs || 0) +
      ' · 表格 ' + Number(stats.tables || 0) +
      ' · 图片 ' + Number(stats.images || 0),
    byRule: byRule.slice(0, 8),
    issues: issues,
    issueTotal: rawIssues.length,
    moreCount: Math.max(0, rawIssues.length - MAX_SHOW_ISSUES)
  };
}

Page({
  data: {
    code: '',
    loading: true,
    errorText: '',
    collection: null,
    rulesetName: '',
    submissions: [],
    subListHeight: 0,
    student: '',
    submitting: false,
    reviewing: false,
    review: null,
    topRules: [],
    reviewAt: '',
    report: null,
    reportLoadingId: '',
    shareTitle: '文档接龙'
  },

  onLoad: function (options) {
    var code = options && options.code ? String(options.code) : '';
    code = code.trim();
    this.hasLoaded = false;
    this.setData({
      code: code,
      student: getApp().readNickname()
    });
    if (!code) {
      this.setData({
        loading: false,
        errorText: '缺少接龙码：请从微信聊天里的接龙卡片进入，或先回首页选择一个接龙。'
      });
      return;
    }
    this.loadDetail(false);
  },

  onShow: function () {
    var nickname = getApp().readNickname();
    if (!this.data.student && nickname) {
      this.setData({ student: nickname });
    }
    // 从分享卡片/H5 返回时静默刷新一次，拿到别人新提交的文档
    if (this.hasLoaded && !this.data.loading) {
      this.loadDetail(true);
    }
  },

  onPullDownRefresh: function () {
    this.loadDetail(true).then(function () {
      wx.stopPullDownRefresh();
    });
  },

  // ---- 加载详情 --------------------------------------------------------
  /** silent=true 时不显示整页 loading（下拉刷新/回到页面时静默刷新）。 */
  loadDetail: function (silent) {
    var that = this;
    if (!this.data.code) {
      return Promise.resolve();
    }
    if (!silent) {
      this.setData({ loading: true, errorText: '' });
    }
    return Promise.all([
      api.getCollection(this.data.code),
      api.getProfiles().catch(function () {
        return null; // 规范名只是展示用，失败不影响主流程
      })
    ]).then(function (results) {
      var payload = results[0] || {};
      var collection = payload.collection;
      if (!collection) {
        throw new Error('接龙不存在或可能已被删除');
      }
      var names = {};
      var profilesPayload = results[1];
      if (profilesPayload && Array.isArray(profilesPayload.profiles)) {
        for (var i = 0; i < profilesPayload.profiles.length; i++) {
          var profile = profilesPayload.profiles[i];
          names[profile.id] = profile.name || profile.id;
        }
      }
      var submissions = [];
      var rawSubs = Array.isArray(payload.submissions) ? payload.submissions : [];
      for (var j = 0; j < rawSubs.length; j++) {
        submissions.push(formatSubmission(rawSubs[j]));
      }
      that.setData({
        loading: false,
        errorText: '',
        collection: formatCollectionDetail(collection, names),
        rulesetName: names[collection.ruleset] || collection.ruleset || '默认规范',
        submissions: submissions,
        subListHeight: Math.max(1, Math.min(submissions.length, 5)) * 260,
        shareTitle: (collection.title || '文档接龙') + '｜接龙码 ' + that.data.code
      });
      that.hasLoaded = true;
      wx.setNavigationBarTitle({ title: collection.title || '接龙详情' });
    }).catch(function (err) {
      var message = (err && err.message) || '加载失败，请稍后重试';
      // 静默刷新失败时保留已渲染内容，不把整页替换成错误
      if (silent && that.data.collection) {
        that.setData({ loading: false });
        wx.showToast({ title: '刷新失败，请稍后重试', icon: 'none' });
        return;
      }
      that.setData({ loading: false, errorText: message });
    });
  },

  onRetry: function () {
    this.loadDetail(false);
  },

  onBackHome: function () {
    var pages = getCurrentPages();
    if (pages.length > 1) {
      wx.navigateBack({ delta: 1 });
      return;
    }
    wx.reLaunch({ url: '/pages/index/index' });
  },

  // ---- 分享 ------------------------------------------------------------
  onShareAppMessage: function () {
    var title = this.data.shareTitle;
    if (this.data.collection) {
      title = this.data.collection.title + '｜文档接龙 ' + this.data.code;
    }
    return {
      title: title,
      path: '/pages/collection/collection?code=' + encodeURIComponent(this.data.code)
    };
  },

  onCopyCode: function () {
    if (!this.data.code) {
      return;
    }
    wx.setClipboardData({ data: this.data.code });
  },

  // ---- 提交文档 --------------------------------------------------------
  onStudentInput: function (event) {
    this.setData({ student: event.detail.value });
  },

  onSubmitTap: function () {
    var that = this;
    if (this.data.submitting) {
      return;
    }
    var student = String(this.data.student || '').trim();
    if (student) {
      this.pickAndSubmit(student);
      return;
    }
    getApp().ensureNickname().then(function (name) {
      that.setData({ student: name });
      that.pickAndSubmit(name);
    }).catch(function (err) {
      that.showError('无法继续', err);
    });
  },

  /** 从微信聊天记录里选 .docx（小程序不能读取手机任意目录）。 */
  pickAndSubmit: function (student) {
    var that = this;
    wx.chooseMessageFile({
      count: 1,
      type: 'file',
      extension: ['docx'],
      success: function (res) {
        var files = (res && res.tempFiles) || [];
        if (!files.length) {
          return;
        }
        var file = files[0];
        var name = file.name || 'document.docx';
        var path = file.path || file.tempFilePath || '';
        if (!/\.docx$/i.test(name)) {
          that.showError('文件格式不支持', new Error('只支持 .docx 文件（.doc 请先用 Word 另存为 .docx）：' + name));
          return;
        }
        if (!path) {
          that.showError('选择文件失败', new Error('没有拿到文件临时路径，请重新选择'));
          return;
        }
        if (file.size && Number(file.size) > api.MAX_UPLOAD_BYTES) {
          that.showError('文件太大', new Error('文件 ' + formatSize(file.size) + '，超过小程序 10MB 上传上限，请压缩图片或删减后重试'));
          return;
        }
        that.doSubmit(student, file, path);
      },
      fail: function (err) {
        var raw = (err && err.errMsg) || '';
        if (raw.indexOf('cancel') >= 0) {
          return; // 用户取消选择，不提示
        }
        that.showError('选择文件失败', new Error(raw || '无法打开聊天文件列表'));
      }
    });
  },

  doSubmit: function (student, file, filePath) {
    var that = this;
    var name = file.name || 'document.docx';
    this.setData({ submitting: true });
    wx.showLoading({ title: '上传检查中', mask: true });
    getApp().saveNickname(student);
    api.submitDocument(this.data.code, student, filePath).then(function (payload) {
      wx.hideLoading();
      var submission = (payload && payload.submission) || {};
      var counts = pickCounts(submission.counts);
      that.setData({
        submitting: false,
        review: null,
        topRules: []
      });
      that.loadDetail(true);
      wx.showModal({
        title: '提交成功',
        content: '文件：' + (submission.filename || name) +
          '\n' + countsText(counts) +
          '\n状态：' + (submission.status || statusFromCounts(counts)) +
          '\n点「一键审核」可查看全部文档汇总。',
        showCancel: false,
        confirmText: '知道了'
      });
    }).catch(function (err) {
      wx.hideLoading();
      that.setData({ submitting: false });
      that.showError('提交失败', err);
    });
  },

  // ---- 一键审核 --------------------------------------------------------
  onReviewTap: function () {
    var that = this;
    if (this.data.reviewing) {
      return;
    }
    if (!this.data.submissions.length) {
      wx.showToast({ title: '还没有人提交文档', icon: 'none' });
      return;
    }
    this.setData({ reviewing: true });
    wx.showLoading({ title: '审核中', mask: true });
    api.reviewCollection(this.data.code).then(function (payload) {
      wx.hideLoading();
      var summary = (payload && payload.summary) || {};
      var counts = pickCounts(summary.counts);
      var submissions = [];
      var rawSubs = (payload && payload.submissions) || [];
      for (var i = 0; i < rawSubs.length; i++) {
        submissions.push(formatSubmission(rawSubs[i]));
      }
      if (!submissions.length) {
        submissions = that.data.submissions;
      }
      that.setData({
        reviewing: false,
        review: {
          total: Number(summary.total || 0),
          checked: Number(summary.checked || 0),
          failed: Number(summary.failed || 0),
          passRateText: formatPassRate(summary.pass_rate),
          error: counts.error,
          warning: counts.warning,
          info: counts.info,
          countsText: countsText(counts)
        },
        topRules: buildTopRules(summary.top_rules),
        reviewAt: nowText(),
        submissions: submissions,
        subListHeight: Math.max(1, Math.min(submissions.length, 5)) * 260
      });
      wx.showToast({ title: '审核完成', icon: 'success' });
    }).catch(function (err) {
      wx.hideLoading();
      that.setData({ reviewing: false });
      that.showError('审核失败', err);
    });
  },

  // ---- 单份报告 --------------------------------------------------------
  onReportTap: function (event) {
    var that = this;
    var submissionId = event.currentTarget.dataset.id;
    if (!submissionId || this.data.reportLoadingId) {
      return;
    }
    this.setData({ reportLoadingId: submissionId });
    wx.showLoading({ title: '读取报告中', mask: true });
    api.getSubmissionReport(this.data.code, submissionId).then(function (payload) {
      wx.hideLoading();
      that.setData({
        reportLoadingId: '',
        report: normalizeReport(payload, submissionId)
      });
      var scroll = function () {
        wx.pageScrollTo({
          selector: '#report-card',
          duration: 300,
          fail: function () {}
        });
      };
      if (wx.nextTick) {
        wx.nextTick(scroll);
      } else {
        setTimeout(scroll, 100);
      }
    }).catch(function (err) {
      wx.hideLoading();
      that.setData({ reportLoadingId: '' });
      that.showError('报告加载失败', err);
    });
  },

  onCloseReport: function () {
    this.setData({ report: null });
  },

  // ---- 提示 ------------------------------------------------------------
  showError: function (title, err) {
    wx.showModal({
      title: title || '出错了',
      content: (err && err.message) || '未知错误，请稍后重试',
      showCancel: false,
      confirmText: '知道了'
    });
  }
});
