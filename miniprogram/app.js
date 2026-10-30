/**
 * tas 文档接龙小程序入口。
 *
 * 只有一个「文档接龙」界面：任何人可发起接龙；同学从分享卡片进入同一接龙，
 * 从微信聊天记录里选 .docx 提交；任何人可点「一键审核」调用后端批量检查。
 * 不做账号登录，昵称只保存在本机 Storage（key: "nickname"）。
 */
App({
  globalData: {
    nickname: ''
  },

  onLaunch: function () {
    this.globalData.nickname = this.readNickname();
  },

  /** 读取本机缓存的昵称（可能为空字符串）。 */
  readNickname: function () {
    try {
      var saved = wx.getStorageSync('nickname');
      return typeof saved === 'string' ? saved.trim() : '';
    } catch (err) {
      return '';
    }
  },

  /** 保存昵称到本机缓存并返回规范化后的值。 */
  saveNickname: function (name) {
    var value = String(name || '').trim();
    if (!value) {
      return '';
    }
    try {
      wx.setStorageSync('nickname', value);
    } catch (err) {
      // 存储失败不影响本次使用（例如存储空间已满）
    }
    this.globalData.nickname = value;
    return value;
  },

  /**
   * 确保有昵称：有缓存直接返回；没有则弹 wx.showModal（editable 输入框）让用户填写。
   * 返回 Promise<string>，取消输入时 reject（页面负责提示）。
   */
  ensureNickname: function () {
    var that = this;
    var cached = this.readNickname();
    if (cached) {
      this.globalData.nickname = cached;
      return Promise.resolve(cached);
    }
    return new Promise(function (resolve, reject) {
      wx.showModal({
        title: '先告诉我你的名字',
        content: '',
        editable: true,
        placeholderText: '输入姓名（只保存在本机，用于接龙显示）',
        confirmText: '保存',
        cancelText: '取消',
        success: function (res) {
          if (!res.confirm) {
            reject(new Error('没有填写姓名，已取消'));
            return;
          }
          var name = that.saveNickname(res.content);
          if (!name) {
            reject(new Error('姓名不能为空'));
            return;
          }
          resolve(name);
        },
        fail: function () {
          reject(new Error('无法打开姓名输入框，请重试'));
        }
      });
    });
  }
});
