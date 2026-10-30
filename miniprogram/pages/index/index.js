var api = require('../../utils/api');

function pad2(value) {
  return value < 10 ? '0' + value : String(value);
}

function todayString() {
  var now = new Date();
  return now.getFullYear() + '-' + pad2(now.getMonth() + 1) + '-' + pad2(now.getDate());
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

/** 把接口返回的 collection 转成页面展示字段（接口字段见 utils/api.js 契约注释）。 */
function formatCollection(item, rulesetNames) {
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
    dueDate: dueDate,
    dueText: dueDate ? '截止 ' + dueDate : '不限截止',
    duePast: !!dueDate && dueDate < todayString(),
    createdAt: formatTime(item.created_at),
    submissionCount: count,
    submissionText: count + ' 人已提交'
  };
}

Page({
  data: {
    loading: true,
    errorText: '',
    collections: [],
    listHeight: 0,
    profiles: [],
    profileLabels: [],
    profileIndex: 0,
    selectedProfileLabel: '',
    defaultProfileId: '',
    dueToday: '',
    createOpen: false,
    creating: false,
    nickname: '',
    form: {
      title: '',
      owner: '',
      ruleset: '',
      note: '',
      due: ''
    }
  },

  onLoad: function () {
    this.rawCollections = [];
    this.setData({ dueToday: todayString() });
    this.loadAll(false);
  },

  onShow: function () {
    var nickname = getApp().readNickname();
    var patch = { nickname: nickname };
    if (!this.data.form.owner && nickname) {
      patch['form.owner'] = nickname;
    }
    this.setData(patch);
  },

  onPullDownRefresh: function () {
    this.loadAll(true).then(function () {
      wx.stopPullDownRefresh();
    });
  },

  // ---- 数据加载 --------------------------------------------------------
  /** silent=true 时不显示整页 loading（用于下拉刷新、创建后静默刷新）。 */
  loadAll: function (silent) {
    var that = this;
    if (!silent) {
      this.setData({ loading: true, errorText: '' });
    }
    return Promise.all([
      api.listCollections(),
      api.getProfiles().catch(function () {
        return null; // 规范接口失败不阻塞列表；发起表单会提示重新加载
      })
    ]).then(function (results) {
      var listPayload = results[0] || {};
      var raw = Array.isArray(listPayload.collections) ? listPayload.collections : [];
      that.applyProfiles(results[1]);
      that.rawCollections = raw;
      that.applyCollections();
      that.setData({ loading: false, errorText: '' });
    }).catch(function (err) {
      var message = (err && err.message) || '加载失败，请稍后重试';
      // 静默刷新失败时保留已有列表，不打断用户
      if (silent && that.data.collections.length) {
        that.setData({ loading: false });
        wx.showToast({ title: '刷新失败，请稍后重试', icon: 'none' });
        return;
      }
      that.setData({ loading: false, errorText: message });
    });
  },

  applyProfiles: function (payload) {
    if (!payload || !Array.isArray(payload.profiles) || payload.profiles.length === 0) {
      return;
    }
    var profiles = payload.profiles;
    var labels = [];
    var names = {};
    for (var i = 0; i < profiles.length; i++) {
      var label = profiles[i].name || profiles[i].id;
      labels.push(label);
      names[profiles[i].id] = label;
    }
    // 优先保留已选规范，否则用后端给的默认规范
    var ruleset = this.data.form.ruleset;
    var index = -1;
    var k = 0;
    for (k = 0; k < profiles.length; k++) {
      if (profiles[k].id === ruleset) {
        index = k;
        break;
      }
    }
    if (index < 0) {
      var preferred = payload.default || profiles[0].id;
      index = 0;
      ruleset = profiles[0].id;
      for (k = 0; k < profiles.length; k++) {
        if (profiles[k].id === preferred) {
          index = k;
          ruleset = profiles[k].id;
          break;
        }
      }
    }
    this.setData({
      profiles: profiles,
      profileLabels: labels,
      profileIndex: index,
      selectedProfileLabel: labels[index] || '',
      defaultProfileId: payload.default || profiles[0].id,
      rulesetNames: names,
      'form.ruleset': ruleset
    });
  },

  applyCollections: function () {
    var names = this.data.rulesetNames || {};
    var raw = this.rawCollections || [];
    var collections = [];
    for (var i = 0; i < raw.length; i++) {
      collections.push(formatCollection(raw[i], names));
    }
    // scroll-view 需要固定高度：按条数估高，最多展示 6 条后内部滚动
    var visible = Math.max(1, Math.min(collections.length, 6));
    this.setData({
      collections: collections,
      listHeight: visible * 270,
      createOpen: this.data.createOpen || collections.length === 0
    });
  },

  // ---- 表单交互 --------------------------------------------------------
  onToggleCreate: function () {
    this.setData({ createOpen: !this.data.createOpen });
  },

  onFormInput: function (event) {
    var field = event.currentTarget.dataset.field;
    if (!field) {
      return;
    }
    var patch = {};
    patch['form.' + field] = event.detail.value;
    this.setData(patch);
  },

  onProfileChange: function (event) {
    var index = Number(event.detail.value);
    var profile = this.data.profiles[index];
    if (!profile) {
      return;
    }
    this.setData({
      profileIndex: index,
      selectedProfileLabel: this.data.profileLabels[index] || profile.name || profile.id,
      'form.ruleset': profile.id
    });
  },

  onDueChange: function (event) {
    this.setData({ 'form.due': event.detail.value || '' });
  },

  onClearDue: function () {
    this.setData({ 'form.due': '' });
  },

  // ---- 发起接龙 --------------------------------------------------------
  onCreateTap: function () {
    var that = this;
    if (this.data.creating) {
      return;
    }
    var form = this.data.form;
    var title = String(form.title || '').trim();
    if (!title) {
      this.toast('请填写接龙标题');
      return;
    }
    var ruleset = form.ruleset || this.data.defaultProfileId;
    if (!ruleset) {
      this.toast('规范列表还没加载好，请下拉刷新重试');
      return;
    }
    var owner = String(form.owner || '').trim();
    if (owner) {
      this.doCreate(title, owner, ruleset);
      return;
    }
    getApp().ensureNickname().then(function (name) {
      that.setData({ 'form.owner': name });
      that.doCreate(title, name, ruleset);
    }).catch(function (err) {
      that.showError('无法继续', err);
    });
  },

  doCreate: function (title, owner, ruleset) {
    var that = this;
    var form = this.data.form;
    this.setData({ creating: true });
    api.createCollection({
      title: title,
      owner: owner,
      ruleset: ruleset,
      note: String(form.note || '').trim(),
      due: String(form.due || '').trim()
    }).then(function (payload) {
      var collection = (payload && payload.collection) || {};
      var code = collection.code || '';
      getApp().saveNickname(owner);
      that.setData({
        creating: false,
        createOpen: false,
        'form.title': '',
        'form.note': '',
        'form.due': ''
      });
      that.loadAll(true);
      if (!code) {
        that.showError('创建成功但未拿到接龙码', new Error('后端响应缺少 collection.code，请检查后端版本'));
        return;
      }
      wx.showModal({
        title: '接龙创建成功',
        content: '接龙码：' + code + '\n把接龙分享给同学，大家从聊天卡片进来，就能提交 .docx。',
        confirmText: '复制并查看',
        cancelText: '留在此页',
        success: function (res) {
          if (res.confirm) {
            wx.setClipboardData({ data: code });
            wx.navigateTo({
              url: '/pages/collection/collection?code=' + encodeURIComponent(code)
            });
          }
        }
      });
    }).catch(function (err) {
      that.setData({ creating: false });
      that.showError('创建失败', err);
    });
  },

  // ---- 列表交互 --------------------------------------------------------
  onCollectionTap: function (event) {
    var code = event.currentTarget.dataset.code;
    if (!code) {
      return;
    }
    wx.navigateTo({
      url: '/pages/collection/collection?code=' + encodeURIComponent(code)
    });
  },

  onRetry: function () {
    this.loadAll(false);
  },

  /**
   * 网络自检：真机连不上后端时点这里。
   * 显示 BASE_URL、手机网络类型、请求耗时、服务端看到的手机 IP。
   */
  onDiagnoseTap: function () {
    var that = this;
    var network = '未知';
    var started = 0;

    /** 运行时改后端地址（存本机），不用改代码/重新编译。 */
    var changeBaseUrl = function () {
      wx.showModal({
        title: '设置后端地址',
        editable: true,
        placeholderText: 'http://电脑局域网IP:8765',
        content: api.getBaseUrl(),
        confirmText: '保存并重试',
        cancelText: '取消',
        success: function (res) {
          if (!res.confirm) {
            return;
          }
          try {
            api.setBaseUrl(res.content);
          } catch (err) {
            that.toast((err && err.message) || '地址不合法');
            return;
          }
          that.toast('已保存，正在重试');
          that.onDiagnoseTap();
        }
      });
    };

    var run = function () {
      wx.showLoading({ title: '检测中', mask: true });
      started = Date.now();
      api.diag().then(function (payload) {
        wx.hideLoading();
        wx.showModal({
          title: '网络正常',
          content: '后端可达：' + api.getBaseUrl() +
            '\n耗时：' + (Date.now() - started) + ' ms' +
            '\n手机网络：' + network +
            '\n电脑看到的手机 IP：' + ((payload && payload.client) || '未知') +
            '\n服务端监听：' + ((payload && payload.server_host) || '?') + ':' + ((payload && payload.server_port) || '?'),
          showCancel: true,
          cancelText: '修改地址',
          confirmText: '知道了',
          success: function (res) {
            if (res.cancel) {
              changeBaseUrl();
            }
          }
        });
      }).catch(function (err) {
        wx.hideLoading();
        wx.showModal({
          title: '网络不通',
          content: '后端地址：' + api.getBaseUrl() +
            '\n错误：' + ((err && err.message) || err) +
            '\n\n点「修改地址」可改后端 IP；也可按顺序检查：' +
            '\n1) 手机和电脑连同一个 Wi-Fi（不要用 4G/校园访客网）' +
            '\n2) 电脑后端用 --host 0.0.0.0 启动，防火墙放行 8765' +
            '\n3) 手机浏览器打开 当前地址 + /api/diag 验证',
          showCancel: true,
          confirmText: '修改地址',
          cancelText: '知道了',
          success: function (res) {
            if (res.confirm) {
              changeBaseUrl();
            }
          }
        });
      });
    };

    wx.getNetworkType({
      success: function (res) {
        network = res.networkType || '未知';
      },
      complete: run
    });
  },

  // ---- 提示 ------------------------------------------------------------
  toast: function (title) {
    wx.showToast({ title: title, icon: 'none', duration: 2200 });
  },

  showError: function (title, err) {
    wx.showModal({
      title: title || '出错了',
      content: (err && err.message) || '未知错误，请稍后重试',
      showCancel: false,
      confirmText: '知道了'
    });
  }
});
