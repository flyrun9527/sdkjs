# ScrollTo 方法场景支持确认

## ✅ 完全支持的场景

### 1. 表格导航 — 滚动到文档中的表格
**支持方式**: ✅ 通过 `ApiTable.GetRange()` 获取 Range，然后调用 `ScrollTo()`

```javascript
window.connector.callCommand(function () {
  var oDocument = Api.GetDocument();
  var tables = oDocument.GetAllTables();
  
  if (tables.length > 0) {
    // 方式1：通过表格的 GetRange 方法
    var tableRange = tables[0].GetRange();
    if (tableRange) {
      tableRange.ScrollTo(true); // ✅ 支持
    }
    
    // 方式2：通过表格第一个单元格的 Range
    var firstCell = tables[0].GetRow(0).GetCell(0);
    var cellContent = firstCell.GetContent();
    var cellRange = cellContent.GetRange();
    if (cellRange) {
      cellRange.ScrollTo(true); // ✅ 支持
    }
  }
}, function (data) {});
```

### 2. 搜索结果批量导航 — 依次浏览所有搜索结果
**支持方式**: ✅ 通过 `GetRangeBySelect()` 获取搜索结果 Range

```javascript
function navigateSearchResults(searchTerm) {
  window.connector.callCommand(function () {
    var oDocument = Api.GetDocument();
    oDocument.SearchAndReplace({ searchString: searchTerm });
    var searchRange = oDocument.GetRangeBySelect();
    if (searchRange) {
      searchRange.ScrollTo(true); // ✅ 支持
    }
  }, function (data) {});
}
```

### 3. 目录（TOC）导航 — 通过目录项跳转
**支持方式**: ✅ 通过标题的 `GetRange()` 方法

```javascript
window.connector.callCommand(function () {
  var oDocument = Api.GetDocument();
  var headings = oDocument.GetAllHeadingParagraphs();
  
  if (headings.length > 0) {
    var headingRange = headings[0].GetRange();
    if (headingRange) {
      headingRange.ScrollTo(true); // ✅ 支持
    }
  }
}, function (data) {});
```

### 4. 内容控件导航 — 导航到表单控件
**支持方式**: ✅ 通过内容控件的 `GetRange()` 方法

```javascript
window.connector.callCommand(function () {
  var oDocument = Api.GetDocument();
  var contentControls = oDocument.GetAllContentControls();
  
  if (contentControls.length > 0) {
    var cc = contentControls[0];
    var ccRange = cc.GetRange();
    if (ccRange) {
      ccRange.ScrollTo(true); // ✅ 支持
    }
  }
}, function (data) {});
```

### 5. 交叉引用导航 — 通过引用跳转
**支持方式**: ✅ 如果引用指向书签或标题，可以通过它们的 Range

```javascript
// 引用指向书签
window.connector.callCommand(function () {
  var oDocument = Api.GetDocument();
  var bookmarkRange = oDocument.GetBookmarkRange("ref_target");
  if (bookmarkRange) {
    bookmarkRange.ScrollTo(true); // ✅ 支持
  }
}, function (data) {});

// 引用指向标题
window.connector.callCommand(function () {
  var oDocument = Api.GetDocument();
  var headings = oDocument.GetAllHeadingParagraphs();
  var targetHeading = headings.find(function(h) {
    return h.GetText().indexOf("目标标题") !== -1;
  });
  if (targetHeading) {
    var headingRange = targetHeading.GetRange();
    if (headingRange) {
      headingRange.ScrollTo(true); // ✅ 支持
    }
  }
}, function (data) {});
```

## ⚠️ 部分支持/需要特殊处理的场景

### 6. 图片/形状导航 — 导航到图片和绘图对象
**支持方式**: ⚠️ 图片/形状可能没有直接的 Range，需要通过其他方式

```javascript
window.connector.callCommand(function () {
  var oDocument = Api.GetDocument();
  var oContent = oDocument.GetContent();
  var images = oContent.GetAllImages();
  
  if (images.length > 0) {
    var image = images[0];
    var imageId = image.GetId();
    
    // 方式1：如果图片在段落中，可以通过段落 Range
    // 注意：需要找到包含图片的段落
    
    // 方式2：使用 SelectDrawingObject（会自动滚动）
    // connector.executeMethod("SelectDrawingObject", [imageId]);
    
    // 方式3：如果图片有对应的 Range（需要检查 API）
    // var imageRange = image.GetRange(); // 可能不存在
    // if (imageRange) {
    //   imageRange.ScrollTo(true);
    // }
  }
}, function (data) {});
```

**建议**: 对于图片/形状，优先使用 `SelectDrawingObject` 方法，它会自动滚动。

### 7. 脚注/尾注导航 — 跳转到脚注位置
**支持方式**: ⚠️ 需要检查脚注/尾注是否有 GetRange 方法

```javascript
window.connector.callCommand(function () {
  var oDocument = Api.GetDocument();
  
  // 方式1：如果脚注有对应的段落，可以通过段落 Range
  // var footnotes = oDocument.GetFootNotesFirstParagraphs();
  // if (footnotes && footnotes.length > 0) {
  //   var footnotePara = footnotes[0];
  //   var footnoteRange = footnotePara.GetRange();
  //   if (footnoteRange) {
  //     footnoteRange.ScrollTo(true);
  //   }
  // }
  
  // 方式2：通过脚注引用找到对应的脚注段落
  // 需要找到文档中引用脚注的位置，然后跳转到脚注区域
}, function (data) {});
```

**建议**: 需要进一步检查 OnlyOffice API 是否提供脚注/尾注的 Range 获取方法。

### 8. 评论/批注导航 — 依次浏览所有评论
**支持方式**: ❌ **ScrollTo 不支持** — ApiComment 没有 GetRange 方法，需要使用 SelectComment

**原因分析**：
- `ApiComment` 类没有提供 `GetRange()` 方法
- 批注内部使用 `RangeStart` 和 `RangeEnd`（ID 字符串），不是 Range 对象
- 虽然可以通过这些 ID 找到对应的段落位置，但 OnlyOffice API 没有提供将批注转换为 Range 的方法

```javascript
// ❌ 不支持：ApiComment 没有 GetRange 方法
// var comment = oDocument.GetCommentById("commentId");
// var commentRange = comment.GetRange(); // 此方法不存在
// commentRange.ScrollTo(true); // 无法执行

// ✅ 推荐方式1：使用 executeMethod 调用 asc_selectComment（会自动滚动）
// 注意：方法名必须是 "asc_selectComment"，只接受一个参数 commentId
window.connector.executeMethod("asc_selectComment", [commentId], function(result) {
  console.log("已滚动到批注:", commentId, result);
});

// ✅ 推荐方式2：使用 callCommand 调用 API
window.connector.callCommand(function () {
  var oDocument = Api.GetDocument();
  var comment = oDocument.GetCommentById(commentId);
  if (comment) {
    // 通过内部方法选择批注（会自动滚动）
    // 注意：这里不能直接调用 SelectComment，需要通过其他方式
    // 建议使用方式1的 executeMethod
  }
}, function (data) {});
```

**建议**: 对于批注，必须使用 `executeMethod("asc_selectComment", [commentId])` 方法。这是唯一支持的方式。

## ⚠️ 常见问题排查

### 问题：executeMethod('SelectComment', [commentId, true]) 无效，无法跳转批注

**原因分析**：
1. ❌ **方法名错误**：应该是 `"asc_selectComment"` 而不是 `"SelectComment"`
2. ❌ **参数错误**：`asc_selectComment` 只接受**一个参数** `commentId`，不接受第二个参数

**正确写法**：
```javascript
// ✅ 正确：方法名是 "asc_selectComment"，只传一个参数 commentId
this.connector.executeMethod('asc_selectComment', [commentId], (result) => {
  console.log('已跳转到批注:', commentId, result);
  this.$message.success(`已跳转到第 ${currentIndex + 1}/${totalCount} 个批注`);
});

// ❌ 错误示例1：方法名错误
this.connector.executeMethod('SelectComment', [commentId, true], callback); // 方法不存在

// ❌ 错误示例2：参数错误（传了两个参数）
this.connector.executeMethod('asc_selectComment', [commentId, true], callback); // 第二个参数会被忽略
```

**代码实现**：
```javascript
// 批注导航示例
goToNextComment() {
  if (this.comments.length === 0) {
    this.$message.warning('文档中没有批注');
    return;
  }
  
  this.currentIndex = (this.currentIndex + 1) % this.comments.length;
  const comment = this.comments[this.currentIndex];
  const commentId = comment.GetId(); // 获取批注 ID
  
  // ✅ 正确调用方式
  this.connector.executeMethod('asc_selectComment', [commentId], (result) => {
    console.log('已跳转到批注:', commentId, result);
    this.$message.success(`已跳转到第 ${this.currentIndex + 1}/${this.comments.length} 个批注`);
  });
}
```

**其他可能的问题**：
1. **commentId 格式错误**：确保 `commentId` 是字符串类型，且是有效的批注 ID
2. **connector 未初始化**：确保 `this.connector` 已经正确初始化
3. **文档未加载完成**：确保文档已完全加载后再调用此方法

### 问题：Api.MoveToNextReviewChange is not a function

**错误信息**：
```
TypeError: Api.MoveToNextReviewChange is not a function
```

**原因分析**：
1. **`MoveToNextReviewChange` 是插件方法**：实际方法名是 `pluginMethod_MoveToNextReviewChange`
2. **`Api` 对象中没有直接暴露这个方法**：`Api` 对象只包含标准的 API 方法，不包含插件方法
3. **插件方法只在插件上下文中可用**：即使使用 `callCommand`，`Api` 对象也不包含插件方法

**错误示例**：
```javascript
// ❌ 错误：Api.MoveToNextReviewChange 不存在
this.connector.callCommand(() => {
  Api.MoveToNextReviewChange(true); // TypeError: Api.MoveToNextReviewChange is not a function
}, callback);

// ❌ 错误：即使使用 false 也会报错
this.connector.callCommand(() => {
  Api.MoveToNextReviewChange(false); // TypeError: Api.MoveToNextReviewChange is not a function
}, callback);
```

**正确写法**：

**方案1：使用 executeMethod 调用底层方法（推荐）**
```javascript
// ✅ 正确：下一个修订
this.connector.executeMethod('asc_GetNextRevisionsChange', [], (result) => {
  console.log('已跳转到下一个修订', result);
  this.$message.success('已跳转到下一个修订');
});

// ✅ 正确：上一个修订
this.connector.executeMethod('asc_GetPrevRevisionsChange', [], (result) => {
  console.log('已跳转到上一个修订', result);
  this.$message.success('已跳转到上一个修订');
});
```

**方案2：使用 callCommand 调用底层方法**
```javascript
// ✅ 正确：下一个修订
this.connector.callCommand(function () {
  this.asc_GetNextRevisionsChange();
}, function (result) {
  console.log('已跳转到下一个修订', result);
});

// ✅ 正确：上一个修订
this.connector.callCommand(function () {
  this.asc_GetPrevRevisionsChange();
}, function (result) {
  console.log('已跳转到上一个修订', result);
});
```

**关键点**：
- **`Api.MoveToNextReviewChange` 不存在**：`Api` 对象中没有这个方法
- **必须使用底层方法**：`asc_GetNextRevisionsChange` 和 `asc_GetPrevRevisionsChange`
- **推荐使用 `executeMethod`**：更简洁，不需要在 `callCommand` 中使用 `this`

**完整实现示例**：
```javascript
// 修订导航管理器
const RevisionNavigator = {
  // 下一个修订
  goToNext() {
    this.connector.executeMethod('asc_GetNextRevisionsChange', [], (result) => {
      if (result) {
        console.log('已跳转到下一个修订');
        this.$message.success('已跳转到下一个修订');
      } else {
        this.$message.info('没有更多修订');
      }
    });
  },
  
  // 上一个修订
  goToPrevious() {
    this.connector.executeMethod('asc_GetPrevRevisionsChange', [], (result) => {
      if (result) {
        console.log('已跳转到上一个修订');
        this.$message.success('已跳转到上一个修订');
      } else {
        this.$message.info('没有更多修订');
      }
    });
  }
};
```

**关键点**：
1. 方法名：`"asc_GetNextRevisionsChange"` 和 `"asc_GetPrevRevisionsChange"`（带 `asc_` 前缀）
2. 参数：这两个方法不需要参数，传空数组 `[]`
3. 返回值：方法可能返回 `null` 或 `undefined` 表示没有更多修订
4. 自动滚动：这些方法内部会自动滚动到修订位置

### 问题：executeMethod('MoveToNextReviewChange', [false], callback) 跳转的是下一个修订而不是上一个

**问题描述**：
通过 postMessage 传递参数：
```javascript
{
  methodName: "MoveToNextReviewChange",
  args: [false],
  type: "method",
  subType: "connector"
}
```
但是跳转的是下一个修订，而不是上一个修订。

**原因分析**：
1. **参数传递问题**：虽然 `callMethodInternal` 会正确添加 `pluginMethod_` 前缀并调用方法，但可能存在以下问题：
   - JSON 序列化/反序列化过程中，`false` 可能被错误处理
   - 参数数组的第一个元素可能不是预期的 `false`
   - `pluginMethod_MoveToNextReviewChange` 的条件判断可能有问题

2. **方法实现逻辑**：
   ```javascript
   Api.prototype["pluginMethod_MoveToNextReviewChange"] = function(isForward)
   {
       if (undefined !== isForward && !isForward)
           this.asc_GetPrevRevisionsChange();  // false -> 上一个
       else
           this.asc_GetNextRevisionsChange();   // true 或 undefined -> 下一个
   };
   ```
   如果 `isForward` 不是严格的 `false`，就会走 `else` 分支，跳转到下一个。

**解决方案**：
**推荐：直接使用底层方法**（最可靠的方式）

```javascript
// ✅ 正确：直接使用底层方法
// 下一个修订
this.connector.executeMethod('asc_GetNextRevisionsChange', [], (result) => {
  console.log('已跳转到下一个修订', result);
});

// 上一个修订
this.connector.executeMethod('asc_GetPrevRevisionsChange', [], (result) => {
  console.log('已跳转到上一个修订', result);
});
```

**如果必须使用 MoveToNextReviewChange**：
1. **检查参数传递**：确保 `args` 数组的第一个元素是严格的 `false`（不是字符串 `"false"` 或其他值）
2. **调试方法**：在 `pluginMethod_MoveToNextReviewChange` 中添加日志，检查实际接收到的参数值
3. **使用 callCommand**：如果 `executeMethod` 有问题，可以尝试使用 `callCommand`：
   ```javascript
   this.connector.callCommand(function () {
     Api.MoveToNextReviewChange(false);  // 上一个
   }, function (result) {
     console.log('已跳转到上一个修订', result);
   });
   ```

**关键点**：
- **最可靠的方式**：使用底层方法 `asc_GetNextRevisionsChange` 和 `asc_GetPrevRevisionsChange`
- 这两个底层方法**不需要参数**，传空数组 `[]`
- 如果使用插件方法，确保参数类型正确（`false` 而不是 `"false"` 或其他值）

### 问题：executeMethod('asc_GetNextRevisionsChange', [], callback) 没效果，未滚动

**可能的原因**：
1. **文档中没有修订**：文档可能没有启用修订跟踪，或者所有修订已被接受/拒绝
2. **修订跟踪未启用**：需要先启用修订跟踪功能
3. **方法调用成功但未找到修订**：可能已经到达最后一个修订
4. **需要检查文档状态**：确保文档已加载完成

**排查步骤**：
```javascript
// 1. 检查文档是否有修订跟踪功能
this.connector.callCommand(function () {
  var oDocument = Api.GetDocument();
  var isTrackRevisions = oDocument.IsTrackRevisions();
  console.log('修订跟踪是否启用:', isTrackRevisions);
  
  // 如果没有启用，可以启用它
  // oDocument.SetTrackRevisions(true);
}, function (data) {});

// 2. 检查返回值，判断是否找到修订
this.connector.executeMethod('asc_GetNextRevisionsChange', [], (result) => {
  console.log('返回值:', result);
  if (result === null || result === undefined) {
    this.$message.warning('没有找到更多修订，可能文档中没有修订或已到达最后一个');
  } else {
    this.$message.success('已跳转到下一个修订');
  }
});
```

**完整实现（带错误处理）**：
```javascript
// 修订导航管理器（带错误处理）
const RevisionNavigator = {
  // 下一个修订
  goToNext() {
    // 先检查是否有修订跟踪
    this.connector.callCommand(function () {
      var oDocument = Api.GetDocument();
      if (!oDocument.IsTrackRevisions()) {
        console.warn('修订跟踪未启用');
        return;
      }
    }, (data) => {
      // 然后尝试导航
      this.connector.executeMethod('asc_GetNextRevisionsChange', [], (result) => {
        console.log('导航结果:', result);
        if (result === null || result === undefined) {
          this.$message.warning('没有更多修订');
        } else {
          this.$message.success('已跳转到下一个修订');
        }
      });
    });
  },
  
  // 上一个修订
  goToPrevious() {
    this.connector.executeMethod('asc_GetPrevRevisionsChange', [], (result) => {
      console.log('导航结果:', result);
      if (result === null || result === undefined) {
        this.$message.warning('没有更多修订');
      } else {
        this.$message.success('已跳转到上一个修订');
      }
    });
  }
};
```

**替代方案（如果 executeMethod 不工作）**：
如果 `executeMethod` 方式不工作，可以尝试使用 `callCommand` 直接调用内部方法：
```javascript
// 方式1：直接调用（如果支持）
this.connector.callCommand(function () {
  var oDocument = Api.GetDocument();
  // 注意：这个方法可能不在 API 中直接暴露
  // 但可以尝试通过其他方式触发
}, function (data) {});

// 方式2：检查方法是否存在
if (this.connector && typeof this.connector.executeMethod === 'function') {
  this.connector.executeMethod('asc_GetNextRevisionsChange', [], callback);
} else {
  console.error('connector.executeMethod 不可用');
}
```

### 9. 修订导航 — 跳转到修订位置
**支持方式**: ❌ **ScrollTo 不支持** — 修订没有 GetRange 方法，需要使用专门的导航方法

**原因分析**：
- 修订（Revision）对象没有提供 `GetRange()` 方法
- 修订导航使用内部方法 `private_SelectRevisionChange()` 来处理选择和滚动
- OnlyOffice API 没有提供将修订转换为 Range 的方法

```javascript
// ❌ 不支持：修订没有 GetRange 方法
// var revision = ...; // 无法直接获取修订对象
// var revisionRange = revision.GetRange(); // 此方法不存在
// revisionRange.ScrollTo(true); // 无法执行

// ✅ 推荐方式1：使用插件 API（推荐）
window.connector.callCommand(function () {
  Api.MoveToNextReviewChange(true); // true 表示下一个，false 表示上一个
}, function (data) {});

// ✅ 推荐方式2：使用底层 API 方法
window.connector.executeMethod("GetNextRevisionsChange", [], function(result) {
  if (result) {
    console.log("已跳转到下一个修订");
  }
});
window.connector.executeMethod("GetPrevRevisionsChange", [], function(result) {
  if (result) {
    console.log("已跳转到上一个修订");
  }
});
```

**建议**: 对于修订，必须使用 `executeMethod("asc_GetNextRevisionsChange", [])` 或 `executeMethod("asc_GetPrevRevisionsChange", [])` 方法。这些方法会自动处理滚动。

**注意**：`Api.MoveToNextReviewChange()` 是插件 API 方法，只在插件上下文中可用。从前端调用需要使用 `executeMethod`。

## 🔍 批注和修订导航实现

### 1. 批注导航 — 跳转到下一个/上一个批注

#### 方式1：使用 GetAllComments + SelectComment（推荐）

```javascript
// 批注导航管理器
var CommentNavigator = {
  comments: [],
  currentIndex: -1,
  
  // 初始化：获取所有批注
  init: function() {
    window.connector.callCommand(function () {
      var oDocument = Api.GetDocument();
      this.comments = oDocument.GetAllComments();
      
      // 按文档位置排序（可选，确保按顺序导航）
      this.comments.sort(function(a, b) {
        // 这里可以根据需要实现排序逻辑
        return 0;
      });
      
      this.currentIndex = -1;
    }.bind(this), function (data) {});
  },
  
  // 跳转到下一个批注
  goToNext: function() {
    if (this.comments.length === 0) {
      console.log("没有批注");
      return;
    }
    
    this.currentIndex = (this.currentIndex + 1) % this.comments.length;
    var comment = this.comments[this.currentIndex];
    var commentId = comment.GetId();
    
    // 使用 asc_selectComment 跳转并滚动
    // 注意：方法名必须是 "asc_selectComment"，只接受一个参数 commentId
    window.connector.executeMethod("asc_selectComment", [commentId], function(result) {
      console.log("已跳转到批注:", commentId, result);
    });
  },
  
  // 跳转到上一个批注
  goToPrevious: function() {
    if (this.comments.length === 0) {
      console.log("没有批注");
      return;
    }
    
    this.currentIndex = this.currentIndex <= 0 
      ? this.comments.length - 1 
      : this.currentIndex - 1;
    var comment = this.comments[this.currentIndex];
    var commentId = comment.GetId();
    
    // 注意：方法名必须是 "asc_selectComment"，只接受一个参数 commentId
    window.connector.executeMethod("asc_selectComment", [commentId], function(result) {
      console.log("已跳转到批注:", commentId, result);
    });
  },
  
  // 跳转到第一个批注
  goToFirst: function() {
    if (this.comments.length === 0) {
      console.log("没有批注");
      return;
    }
    
    this.currentIndex = 0;
    var comment = this.comments[this.currentIndex];
    var commentId = comment.GetId();
    
    // 注意：方法名必须是 "asc_selectComment"，只接受一个参数 commentId
    window.connector.executeMethod("asc_selectComment", [commentId], function(result) {
      console.log("已跳转到第一个批注:", commentId, result);
    });
  }
};

// 使用示例
CommentNavigator.init();
// 点击"下一个批注"按钮时调用
// CommentNavigator.goToNext();
// 点击"上一个批注"按钮时调用
// CommentNavigator.goToPrevious();
```

#### 方式2：使用 API 方法（通过 callCommand）

```javascript
// 跳转到下一个批注
function navigateToNextComment() {
  window.connector.callCommand(function () {
    var oDocument = Api.GetDocument();
    var comments = oDocument.GetAllComments();
    
    if (comments.length === 0) {
      console.log("文档中没有批注");
      return;
    }
    
    // 获取当前选中的批注（如果有）
    // 注意：OnlyOffice API 可能不直接提供获取当前选中批注的方法
    // 需要自己维护状态或通过其他方式判断
    
    // 假设我们要跳转到第一个批注
    var firstComment = comments[0];
    var commentId = firstComment.GetId();
    
    // 使用 executeMethod 调用 asc_selectComment
    // 注意：方法名必须是 "asc_selectComment"，只接受一个参数 commentId
    window.connector.executeMethod("asc_selectComment", [commentId], function(result) {
      console.log("已跳转到批注", result);
    });
  }, function (data) {});
}
```

### 2. 修订导航 — 跳转到下一个/上一个修订

#### 方式1：使用 executeMethod 调用底层 API（推荐）

```javascript
// 跳转到下一个修订
function navigateToNextRevision() {
  // ✅ 正确：使用 executeMethod 调用底层方法
  window.connector.executeMethod("asc_GetNextRevisionsChange", [], function(result) {
    if (result) {
      console.log("已跳转到下一个修订");
    } else {
      console.log("没有更多修订");
    }
  });
}

// 跳转到上一个修订
function navigateToPreviousRevision() {
  // ✅ 正确：使用 executeMethod 调用底层方法
  window.connector.executeMethod("asc_GetPrevRevisionsChange", [], function(result) {
    if (result) {
      console.log("已跳转到上一个修订");
    } else {
      console.log("没有更多修订");
    }
  });
}
```

**注意**：`Api.MoveToNextReviewChange` 是插件 API 方法，只在插件上下文中可用。从前端调用需要使用 `executeMethod` 调用底层方法。

#### 方式2：使用底层 API 方法

```javascript
// 跳转到下一个修订
function navigateToNextRevision() {
  window.connector.executeMethod("GetNextRevisionsChange", [], function(result) {
    if (result) {
      console.log("已跳转到下一个修订");
    } else {
      console.log("没有更多修订");
    }
  });
}

// 跳转到上一个修订
function navigateToPreviousRevision() {
  window.connector.executeMethod("GetPrevRevisionsChange", [], function(result) {
    if (result) {
      console.log("已跳转到上一个修订");
    } else {
      console.log("没有更多修订");
    }
  });
}
```

### 3. 完整的批注和修订导航组件示例

```javascript
// 完整的导航管理器
var DocumentNavigator = {
  // 批注相关
  commentNavigator: {
    comments: [],
    currentIndex: -1,
    
    init: function() {
      window.connector.callCommand(function () {
        var oDocument = Api.GetDocument();
        this.comments = oDocument.GetAllComments();
        this.currentIndex = -1;
      }.bind(this), function (data) {});
    },
    
    goToNext: function() {
      if (this.comments.length === 0) {
        alert("文档中没有批注");
        return;
      }
      this.currentIndex = (this.currentIndex + 1) % this.comments.length;
      var commentId = this.comments[this.currentIndex].GetId();
      // 注意：方法名必须是 "asc_selectComment"，只接受一个参数 commentId
      window.connector.executeMethod("asc_selectComment", [commentId]);
    },
    
    goToPrevious: function() {
      if (this.comments.length === 0) {
        alert("文档中没有批注");
        return;
      }
      this.currentIndex = this.currentIndex <= 0 
        ? this.comments.length - 1 
        : this.currentIndex - 1;
      var commentId = this.comments[this.currentIndex].GetId();
      // 注意：方法名必须是 "asc_selectComment"，只接受一个参数 commentId
      window.connector.executeMethod("asc_selectComment", [commentId]);
    },
    
    getCurrentComment: function() {
      if (this.currentIndex >= 0 && this.currentIndex < this.comments.length) {
        return this.comments[this.currentIndex];
      }
      return null;
    },
    
    getCommentCount: function() {
      return this.comments.length;
    }
  },
  
  // 修订相关
  revisionNavigator: {
    goToNext: function() {
      // ✅ 使用 executeMethod 调用底层方法
      window.connector.executeMethod("asc_GetNextRevisionsChange", [], function(result) {
        if (result) {
          console.log("已跳转到下一个修订");
        }
      });
    },
    
    goToPrevious: function() {
      // ✅ 使用 executeMethod 调用底层方法
      window.connector.executeMethod("asc_GetPrevRevisionsChange", [], function(result) {
        if (result) {
          console.log("已跳转到上一个修订");
        }
      });
    }
  },
  
  // 初始化
  init: function() {
    this.commentNavigator.init();
  }
};

// 使用示例
DocumentNavigator.init();

// HTML 按钮绑定示例
// <button onclick="DocumentNavigator.commentNavigator.goToNext()">下一个批注</button>
// <button onclick="DocumentNavigator.commentNavigator.goToPrevious()">上一个批注</button>
// <button onclick="DocumentNavigator.revisionNavigator.goToNext()">下一个修订</button>
// <button onclick="DocumentNavigator.revisionNavigator.goToPrevious()">上一个修订</button>
```

### 4. 注意事项

1. **批注导航**：
   - `GetAllComments()` 返回所有批注，需要自己维护当前索引
   - 使用 `executeMethod("asc_selectComment", [commentId])` 来选择和滚动到批注
   - **重要**：方法名必须是 `"asc_selectComment"`（不是 `"SelectComment"`），只接受一个参数 `commentId`
   - `asc_selectComment` 方法内部已经硬编码了滚动参数，会自动滚动到批注位置
   - 批注可能没有按文档顺序排列，如需按顺序导航，需要自己实现排序逻辑

2. **修订导航**：
   - **重要**：`Api.MoveToNextReviewChange()` 是插件 API 方法，只在插件上下文中可用
   - 从前端调用需要使用 `executeMethod` 调用底层方法：
     - `executeMethod("asc_GetNextRevisionsChange", [])` - 下一个修订
     - `executeMethod("asc_GetPrevRevisionsChange", [])` - 上一个修订
   - 这些方法会自动处理滚动和选择
   - 修订导航会自动循环（到达最后一个后回到第一个）

3. **错误处理**：
   - 检查批注/修订是否存在
   - 处理空文档或没有批注/修订的情况
   - 监听导航结果，提供用户反馈

## 📋 总结

| 场景 | 支持状态 | 使用方式 |
|------|---------|---------|
| 表格导航 | ✅ 完全支持 | `table.GetRange().ScrollTo()` |
| 搜索结果导航 | ✅ 完全支持 | `doc.GetRangeBySelect().ScrollTo()` |
| 目录（TOC）导航 | ✅ 完全支持 | `heading.GetRange().ScrollTo()` |
| 内容控件导航 | ✅ 完全支持 | `contentControl.GetRange().ScrollTo()` |
| 交叉引用导航 | ✅ 完全支持 | `bookmarkRange.ScrollTo()` 或 `headingRange.ScrollTo()` |
| 图片/形状导航 | ⚠️ 部分支持 | 使用 `SelectDrawingObject` 或查找包含图片的段落 Range |
| 脚注/尾注导航 | ⚠️ 需要验证 | 可能需要通过脚注段落获取 Range |
| 评论/批注导航 | ❌ 不支持 ScrollTo | 使用 `SelectComment(commentId, true)` |
| 修订导航 | ❌ 不支持 ScrollTo | 使用 `executeMethod("asc_GetNextRevisionsChange", [])` 或 `executeMethod("asc_GetPrevRevisionsChange", [])` |

## 🔧 改进建议

1. **图片/形状**: 如果 OnlyOffice API 支持获取图片/形状的 Range，可以直接使用 `ScrollTo()`。否则使用 `SelectDrawingObject`。

2. **脚注/尾注**: 检查是否有 API 可以获取脚注/尾注的 Range，如果没有，可能需要通过脚注引用位置来定位。

3. **批注**: `ScrollTo()` **不支持**批注导航。`ApiComment` 没有 `GetRange()` 方法，必须使用 `SelectComment(commentId, true)` 作为替代方案。

4. **修订**: `ScrollTo()` **不支持**修订导航。修订对象没有 `GetRange()` 方法，必须使用 `executeMethod("asc_GetNextRevisionsChange", [])` 或 `executeMethod("asc_GetPrevRevisionsChange", [])` 方法。

## ✅ 结论

**当前 `ScrollTo()` 实现支持大部分场景**（5/9 完全支持，2/9 部分支持，2/9 不支持）。

对于不完全支持的场景，都有相应的替代方案：
- 图片/形状 → 使用 `SelectDrawingObject`
- 批注 → 使用 `SelectComment(commentId, true)`（**ScrollTo 不支持**）
- 修订 → 使用 `executeMethod("asc_GetNextRevisionsChange", [])` 或 `executeMethod("asc_GetPrevRevisionsChange", [])`（**ScrollTo 不支持**）
- 脚注/尾注 → 需要进一步验证 API 支持情况

## 📌 ScrollTo 对批注和修订的支持总结

| 功能 | ScrollTo 支持 | 替代方案 |
|------|--------------|---------|
| **批注导航** | ❌ **不支持** | `SelectComment(commentId, true)` |
| **修订导航** | ❌ **不支持** | `executeMethod("asc_GetNextRevisionsChange", [])` 或 `executeMethod("asc_GetPrevRevisionsChange", [])` |

**结论**：
- `ScrollTo()` 方法**不能直接**用于滚动到批注或修订位置
- 批注和修订需要使用各自专门的导航方法，这些方法内部已经实现了滚动功能
- 如果未来 OnlyOffice API 为 `ApiComment` 和修订对象添加 `GetRange()` 方法，则可以使用 `ScrollTo()`




