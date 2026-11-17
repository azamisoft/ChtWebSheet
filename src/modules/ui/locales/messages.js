// Japanese UI source labels are stable keys. User workbook/chat content is excluded by the localizer.
export function localeMessages() {
  return Object.fromEntries(`
ファイル|File|文件
ホーム|Home|开始
挿入|Insert|插入
描画|Draw|绘图
レイアウト|Layout|布局
数式|Formulas|公式
データ|Data|数据
表示|View|视图
ヘルプ|Help|帮助
言語|Language|语言
保存|Save|保存
開く|Open|打开
閉じる|Close|关闭
新規|New|新建
新しいブック|New workbook|新建工作簿
新規作成|Create new|新建
名前を付けて保存|Save as|另存为
ファイルを選択|Choose file|选择文件
Excel ファイルを選択してください。|Choose an Excel file.|请选择 Excel 文件。
Cht WebSheet を読み込み中...|Loading Cht WebSheet…|正在加载 Cht WebSheet…
CWS HTML または Excel ファイルをここにドロップ|Drop a CWS HTML or Excel file here|将 CWS HTML 或 Excel 文件拖放到此处
.html / .xlsx / .xlsm / .xls に対応しています。数式は HyperFormula で計算し、基本スタイル、結合セル、リンク、ドロップダウン リストをできるだけ保持します。|Supports .html / .xlsx / .xlsm / .xls. HyperFormula calculates formulas; formatting, merged cells, links and dropdown lists are preserved where possible.|支持 .html / .xlsx / .xlsm / .xls。使用 HyperFormula 计算公式，并尽可能保留格式、合并单元格、链接及下拉列表。
未読み込み|No file loaded|尚未加载
現在のファイル名|Current file name|当前文件名
ページを再読み込みしても内容は保持されます。|Your work is kept when you reload.|刷新页面后仍会保留内容。
Web版で開く|Open web version|打开网页版
貼り付け|Paste|粘贴
切り取り|Cut|剪切
コピー|Copy|复制
書式のコピー/貼り付け|Format painter|格式刷
クリップボード|Clipboard|剪贴板
フォント|Font|字体
フォントサイズ|Font size|字号
太字|Bold|加粗
斜体|Italic|斜体
下線|Underline|下划线
下線を適用|Apply underline|添加下划线
下線メニューを開く|Underline options|下划线选项
罫線|Borders|边框
罫線を適用|Apply borders|添加边框
罫線メニューを開く|Border options|边框选项
背景色|Fill color|填充颜色
背景色を適用|Apply fill color|应用填充颜色
背景色の色板を開く|Choose fill color|选择填充颜色
文字色|Font color|文字颜色
文字色を適用|Apply font color|应用文字颜色
文字色の色板を開く|Choose font color|选择文字颜色
上揃え|Top align|顶端对齐
上下中央揃え|Middle align|垂直居中
下揃え|Bottom align|底端对齐
左揃え|Align left|左对齐
中央揃え|Center|居中
右揃え|Align right|右对齐
折り返して全体を表示する|Wrap text|自动换行
セルを結合して中央揃え|Merge & center|合并后居中
配置|Alignment|对齐方式
標準|Normal|普通
通貨表示形式|Currency format|货币格式
パーセント スタイル|Percent style|百分比格式
桁区切りスタイル|Comma style|千位分隔符
小数点以下の表示桁数を増やす|Increase decimals|增加小数位数
小数点以下の表示桁数を減らす|Decrease decimals|减少小数位数
数値|Number|数字
文字列|Text|文本
日付|Date|日期
パーセンテージ|Percentage|百分比
通貨|Currency|货币
表示形式|Number format|数字格式
セル|Cells|单元格
削除|Delete|删除
書式|Format|格式
オートSUM|AutoSum|自动求和
フィル|Fill|填充
クリア|Clear|清除
並べ替えとフィルター|Sort & filter|排序和筛选
検索と選択|Find & select|查找和选择
編集|Editing|编辑
ピボット|Pivot|数据透视
テーブル|Table|表格
図|Illustrations|插图
図形|Shapes|形状
画像|Pictures|图片
アイコン|Icons|图标
3D モデル|3D models|3D 模型
スクリーンショット|Screenshot|屏幕截图
グラフ|Charts|图表
縦棒/横棒グラフ|Column / bar chart|柱形图／条形图
折れ線/面グラフ|Line / area chart|折线图／面积图
円/ドーナツ グラフ|Pie / doughnut chart|饼图／圆环图
散布図/バブル|Scatter / bubble|散点图／气泡图
おすすめ|Recommended|推荐
マップ|Maps|地图
ピボットグラフ|Pivot chart|数据透视图
スパークライン|Sparklines|迷你图
折れ線|Line|折线
縦棒|Column|柱形
勝敗|Win/loss|盈亏
フィルター|Filter|筛选
スライサー|Slicer|切片器
タイムライン|Timeline|时间线
リンク|Link|链接
テキスト|Text|文本
テキスト ボックス|Text box|文本框
ヘッダーとフッター|Header & footer|页眉和页脚
ワードアート|WordArt|艺术字
署名欄|Signature line|签名栏
オブジェクト|Object|对象
記号と特殊文字|Symbols|符号与特殊字符
記号|Symbol|符号
描画ツール|Drawing tools|绘图工具
選択|Select|选择
なげなわ選択|Lasso select|套索选择
消しゴム|Eraser|橡皮擦
ペン|Pen|钢笔
鉛筆|Pencil|铅笔
蛍光ペン|Highlighter|荧光笔
インクの削除|Delete ink|删除墨迹
変換|Convert|转换
図形に変換|Convert to shape|转换为形状
テーマ|Themes|主题
配色|Colors|配色
効果|Effects|效果
ページ設定|Page setup|页面设置
余白|Margins|页边距
印刷の向き|Orientation|纸张方向
サイズ|Size|大小
印刷範囲|Print area|打印区域
改ページ|Page breaks|分页符
背景|Background|背景
印刷タイトル|Print titles|打印标题
拡大縮小印刷|Scale to fit|缩放打印
幅|Width|宽度
高さ|Height|高度
拡大/縮小|Scale|缩放
シートのオプション|Sheet options|工作表选项
目盛線|Gridlines|网格线
印刷|Print|打印
見出し|Headings|标题
整列|Align|对齐
前面へ移動|Bring forward|上移一层
背面へ移動|Send backward|下移一层
選択ウィンドウ|Selection pane|选择窗格
グループ化|Group|组合
グループ解除|Ungroup|取消组合
回転|Rotate|旋转
関数の|Function|函数
関数ライブラリ|Function library|函数库
最近使った|Recently used|最近使用
財務|Financial|财务
論理|Logical|逻辑
日付/時刻|Date & time|日期和时间
検索/行列|Lookup & reference|查找与引用
数学/三角|Math & trigonometry|数学与三角
その他|More|更多
定義された名前|Defined names|定义的名称
名前の管理|Name manager|名称管理器
名前の定義|Define name|定义名称
数式で使用|Use in formula|用于公式
名前の貼り付け|Paste names|粘贴名称
選択範囲から作成|Create from selection|根据所选内容创建
ワークシート分析|Formula auditing|公式审核
参照元のトレース|Trace precedents|追踪引用单元格
参照先のトレース|Trace dependents|追踪从属单元格
トレース矢印の削除|Remove arrows|移去箭头
数式の表示|Show formulas|显示公式
エラー チェック|Error checking|错误检查
数式の検証|Evaluate formula|公式求值
ウォッチ ウィンドウ|Watch window|监视窗口
計算方法|Calculation|计算
計算方法の設定|Calculation options|计算选项
再計算実行|Calculate now|立即计算
シート再計算|Calculate sheet|计算工作表
データの取得と変換|Get & transform data|获取和转换数据
データの|Data|数据
取得|Get|获取
ソース|Sources|数据源
最近使用|Recent sources|最近使用的数据源
既存の|Existing|现有
接続|Connections|连接
クエリと|Queries &|查询与
更新|Refresh|刷新
すべて|All|全部
並べ替え|Sort|排序
詳細|Advanced|高级
再適用|Reapply|重新应用
データ ツール|Data tools|数据工具
区切り位置|Text to columns|分列
フラッシュ|Flash|快速
重複の削除|Remove duplicates|删除重复项
入力規則|Data validation|数据验证
統合|Consolidate|合并计算
リレーション|Relation|关系
シップ|ships|关系
モデル|Model|模型
予測|Forecast|预测
分析|Analysis|分析
アウトライン|Outline|分级显示
小計|Subtotal|分类汇总
詳細表示|Show details|显示明细
詳細非表示|Hide details|隐藏明细
シート ビュー|Sheet view|工作表视图
個人用|Personal|个人
ブックの表示|Workbook views|工作簿视图
改ページプレビュー|Page break preview|分页预览
ページレイアウト|Page layout|页面布局
ユーザー設定のビュー|Custom views|自定义视图
ルーラー|Ruler|标尺
数式バー|Formula bar|编辑栏
数式バーを展開|Expand formula bar|展开编辑栏
数式バーの高さ|Formula bar height|编辑栏高度
名前ボックス|Name box|名称框
ズーム|Zoom|缩放
ズームイン|Zoom in|放大
ズームアウト|Zoom out|缩小
ズーム設定|Zoom settings|缩放设置
選択範囲に合わせる|Zoom to selection|缩放到选区
ウィンドウ|Window|窗口
新しいウィンドウ|New window|新建窗口
整列|Arrange|排列
ウィンドウ枠の固定|Freeze panes|冻结窗格
ウィンドウ枠の固定メニュー|Freeze pane options|冻结窗格选项
分割|Split|拆分
表示しない|Hide|隐藏
再表示|Unhide|取消隐藏
並べて比較|View side by side|并排查看
同時にスクロール|Synchronous scrolling|同步滚动
ウィンドウ位置のリセット|Reset window position|重设窗口位置
ウィンドウの切り替え|Switch windows|切换窗口
マクロ|Macros|宏
マクロ メニュー|Macro menu|宏菜单
図形の書式|Shape format|形状格式
図形の挿入|Insert shapes|插入形状
図形のスタイル|Shape styles|形状样式
図形の塗りつぶし|Shape fill|形状填充
図形の枠線|Shape outline|形状轮廓
図形の効果|Shape effects|形状效果
ワードアートのスタイル|WordArt styles|艺术字样式
文字の塗りつぶし|Text fill|文字填充
文字の輪郭|Text outline|文字轮廓
テキストの編集|Edit text|编辑文字
図形メニューを開く|Open shape menu|打开形状菜单
トリミング|Crop|裁剪
サイズのリセット|Reset size|重置大小
塗りつぶし|Fill|填充
バージョン|Version|版本
リボン|Ribbon|功能区
リボンのレイアウト|Ribbon layout|功能区布局
リボン表示オプション|Ribbon display options|功能区显示选项
クラシック リボン|Classic ribbon|经典功能区
1行のリボン|Single-line ribbon|单行功能区
リボンを表示|Show ribbon|显示功能区
常に表示|Always show|始终显示
自動的に非表示にする|Auto-hide|自动隐藏
シート|Sheet|工作表
シートの追加|Add sheet|添加工作表
シート見出し|Sheet tabs|工作表标签
シート見出しのスクロール|Scroll sheet tabs|滚动工作表标签
シート見出しを右へスクロール|Scroll sheet tabs right|向右滚动工作表标签
シート見出しを左へスクロール|Scroll sheet tabs left|向左滚动工作表标签
ドラッグしてシート見出し領域を調整|Drag to resize sheet tabs|拖动调整工作表标签区域
横スクロール バー|Horizontal scrollbar|水平滚动条
横スクロールバー|Horizontal scrollbar|水平滚动条
縦スクロールバー|Vertical scrollbar|垂直滚动条
ワークブック ウィンドウ|Workbook windows|工作簿窗口
選択範囲の集計|Selection summary|选区统计
設定|Settings|设置
オプション|Options|选项
オート|Auto|自动
自動|Automatic|自动
ゼロ値|Zero values|零值
右から左|Right to left|从右到左
アウトライン記号|Outline symbols|分级显示符号
保持|Keep|保留
終了|Exit|退出
全画面表示|Full screen|全屏
表示とズーム|View & zoom|视图与缩放
ページ|Page|页面
ページ余白|Page margins|页边距
1 ページ|One page|单页
2 ページ|Two pages|双页
横|Landscape|横向
縦|Portrait|纵向
枠線|Gridlines|网格线
プレビュー|Preview|预览
から|From|从
または範囲から|Or from range|或从区域
を入手|Get|获取
アドイン|Add-ins|加载项
細線|Thin|细线
太線|Thick|粗线
黒|Black|黑色
赤|Red|红色
青|Blue|蓝色
黄|Yellow|黄色
淡色 - 青|Light blue|浅蓝色
淡色 - 緑|Light green|浅绿色
淡色 - オレンジ|Light orange|浅橙色
濃色 - 紺|Dark blue|深蓝色
キャンセル|Cancel|取消
適用|Apply|应用
確認|Confirm|确认
名前|Name|名称
種類|Type|类型
内容|Content|内容
名前を変更|Rename|重命名
名前を変更...|Rename…|重命名…
移動またはコピー|Move or copy|移动或复制
シートの保護|Protect sheet|保护工作表
シートの削除|Delete sheet|删除工作表
行の高さ|Row height|行高
列の幅|Column width|列宽
行|Row|行
列|Column|列
幅の自動調整|AutoFit width|自动调整宽度
高さの自動調整|AutoFit height|自动调整高度
元に戻す|Undo|撤销
やり直し|Redo|重做
検索|Find|查找
置換|Replace|替换
検索する文字列|Find what|查找内容
置換後の文字列|Replace with|替换为
次を検索|Find next|查找下一个
すべて置換|Replace all|全部替换
検索場所|Within|搜索范围
ブック|Workbook|工作簿
大文字と小文字を区別する|Match case|区分大小写
セル内容が完全に同一であるものを検索する|Match entire cell contents|匹配整个单元格内容
閉じる|Close|关闭
最近使ったファイル|Recent files|最近使用的文件
参照|Browse|浏览
ダウンロード|Downloads|下载
デスクトップ|Desktop|桌面
ドキュメント|Documents|文档
この PC|This PC|此电脑
ブラウザーの既定の保存先|Default browser download location|浏览器默认下载位置
保存先を参照して指定|Browse to choose a location|浏览并选择保存位置
参照して保存先を選択|Browse to select a save location|浏览选择保存位置
ローカル共通ランタイム用|For local shared runtime|供本地共享运行时使用
ファイル名|File name|文件名
ファイルの種類|File type|文件类型
保存先|Save location|保存位置
上書き保存|Save|保存
エクスポート|Export|导出
インポート|Import|导入
情報|Info|信息
印刷プレビュー|Print preview|打印预览
最近使ったブック|Recent workbooks|最近使用的工作簿
空白のブック|Blank workbook|空白工作簿
セッション|Sessions|会话
新規チャット|New chat|新建对话
チャットに戻る|Back to chat|返回对话
ログアウト|Sign out|退出登录
ログイン|Sign in|登录
メールアドレス|Email address|邮箱地址
パスワード|Password|密码
ログインする|Sign in|登录
新規登録|Sign up|注册
アカウントを作成|Create account|创建账号
登録|Register|注册
会話の開始|Conversation started|对话开始
メッセージ|Message|消息
メッセージはありません|No messages yet|暂无消息
送信|Send|发送
送信中|Sending|正在发送
添付ファイルを追加|Add attachment|添加附件
添付ファイル|Attachments|附件
AI を閉じる|Close AI|关闭 AI
AI を開く|Open AI|打开 AI
AI を開く/閉じる|Toggle AI|打开／关闭 AI
サイズ変更|Resize|调整大小
AI 接続設定|AI connection settings|AI 连接设置
API プロバイダー|API provider|API 提供商
モデル ID|Model ID|模型 ID
推論レベル low を送信|Send low reasoning effort|使用 low 推理级别
接続設定を保存|Save connection settings|保存连接设置
保存済みのキーを削除|Delete saved key|删除已保存的密钥
保存済み（変更時のみ入力）|Saved (enter only to replace)|已保存（仅更改时输入）
接続設定を保存しました。|Connection settings saved.|连接设置已保存。
未完了の作業|Incomplete task|未完成的任务
未完了の作業があります|There is an incomplete task|有尚未完成的任务
保存済みの検証結果から再開できます。|Resume from the saved checkpoint.|可从已保存的检查点恢复。
作業を再開|Resume task|恢复任务
再開を終了|Dismiss recovery|结束恢复
押して変更前を確認|Hold to compare before|按住查看修改前
変更前を表示中|Showing before changes|正在显示修改前
未通過・未確認の項目|Failed or unverified checks|未通过或尚未确认的项目
適用する|Apply|应用
キャンセルする|Cancel|取消
変更を適用|Apply changes|应用变更
変更を取り消す|Revert changes|撤销变更
途中結果|Partial results|部分结果
AI接続エラー|AI connection error|AI 连接错误
AI休眠中|AI is offline|AI 处于离线状态
接続を確認中|Checking connection|正在检查连接
ブック内容の処理方法を確認中|Checking how to handle the workbook|正在判断工作簿处理方式
プランを確認中|Reviewing plan|正在确认计划
実行中|Running|执行中
処理中|Processing|处理中
検証中|Validating|验证中
完了|Complete|完成
未完了|Incomplete|未完成
失敗|Failed|失败
停止|Stop|停止
履歴|History|历史记录
検索・記憶の詳細を表示|Show retrieval and memory details|显示检索与记忆详情
開発者設定|Developer settings|开发者设置
契約プラン・週間利用枠|Plan & weekly allowance|套餐与每周额度
個人の API Key に課金されます。|Your personal API key is billed.|费用由您的个人 API Key 账户承担。
CHT の利用枠は消費しません。|Your CHT allowance is not charged.|不扣除 CHT 套餐额度。
選択した API に依頼内容・必要なブック情報・添付ファイルを送信します。料金は個人の API Key に請求されます。CHT の利用枠は消費しません。|Requests, required workbook data and attachments are sent to the selected API. Your personal API key is billed; your CHT allowance is not charged.|请求、必要的工作簿信息及附件将发送至所选 API。费用由个人 API Key 账户承担，不扣除 CHT 套餐额度。
AI サービスに接続できません。接続設定・通信・Harness などの中継サービスを確認して更新してください。|Cannot connect to the AI service. Check connection settings, network and relay services, then refresh.|无法连接 AI 服务。请检查连接设置、网络和中转服务后刷新。
システムプロンプト|System prompt|系统提示词
既定に戻す|Reset to default|恢复默认
保存しました。|Saved.|已保存。
読み込み中...|Loading…|加载中…
アクセス権限|Permissions|访问权限
詳細設定|Advanced settings|高级设置
実行|Run|运行
削除する|Delete|删除
戻る|Back|返回
次へ|Next|下一步
いいえ|No|否
はい|Yes|是
リンク先|Link address|链接地址
アドレス|Address|地址
表示文字列|Display text|显示文本
スクロール バー|Scrollbars|滚动条
言語を選択|Choose language|选择语言
リレーションシップ|Relationships|关系
クエリと接続|Queries & connections|查询与连接
データの取得|Get data|获取数据
フラッシュフィル|Flash fill|快速填充
関数の挿入|Insert function|插入函数
ピボットテーブル|Pivot table|数据透视表
おすすめピボット|Recommended pivots|推荐透视表
アドインを入手|Get add-ins|获取加载项
個人用アドイン|My add-ins|我的加载项
おすすめグラフ|Recommended charts|推荐图表
オートSUM|AutoSum|自动求和
テキストCSV|Text / CSV|文本 / CSV
Webから|From web|从网页
テーブルまたは範囲から|From table / range|从表格或范围
最近使ったソース|Recent sources|最近使用的数据源
既存の接続|Existing connections|现有连接
データの入力規則|Data validation|数据验证
データモデル|Data model|数据模型
What-If分析|What-if analysis|假设分析
予測シート|Forecast sheet|预测工作表
データ分析|Data analysis|数据分析
ページレイアウト|Page layout|页面布局
改ページプレビュー|Page break preview|分页预览
直接入力|Enter directly|直接输入
すべて更新|Refresh all|全部刷新
新しいブックを作成しました。|Created a new workbook.|已创建新工作簿。
表示名|Display name|显示名称
パスワードをお忘れですか|Forgot password?|忘记密码？
AIアシスタントはただいまこちらへ向かっています|The AI assistant is getting ready.|AI 助手正在准备中
もう少しだけお待ちください。|Please wait a moment.|请稍候。
パスワード再設定|Reset password|重置密码
確認コードをメールで送信します。|We will email you a verification code.|我们将向你的邮箱发送验证码。
新しいパスワード|New password|新密码
確認コードを送信|Send verification code|发送验证码
ログインに戻る|Back to sign in|返回登录
確認コード|Verification code|验证码
コードを再送|Resend code|重新发送验证码
メールを変更|Change email|更改邮箱
再試行|Try again|重试
セッションはありません|No conversations yet|暂无会话
添付を削除|Remove attachment|移除附件
復元診断を保存（ブック内容を含む）|Save recovery diagnostics (includes workbook)|保存恢复诊断（包含工作簿内容）
チャット設定|Chat settings|聊天设置
中断した作業の復元|Recover interrupted work|恢复中断的任务
保存した復元診断を、同じブックと元のセッションに読み込みます。|Load saved recovery diagnostics into the same workbook and original conversation.|将已保存的恢复诊断导入同一工作簿和原会话。
復元データを読み込む|Load recovery data|导入恢复数据
パスワード変更|Change password|修改密码
現在のパスワード|Current password|当前密码
パスワードを変更|Change password|修改密码
CWS に変更を反映しています...|Applying changes to CWS…|正在将更改应用到 CWS…
OpenAI 互換 API|OpenAI-compatible API|OpenAI 兼容 API
ベース URL|Base URL|基础 URL
WebSheet 共通ランタイムを読み込み中...|Loading the WebSheet runtime…|正在加载 WebSheet 运行环境…
HTML ワークブックを読み込みました。|Loaded the HTML workbook.|已加载 HTML 工作簿。
開くファイルを選択してください。|Choose a file to open.|请选择要打开的文件。
ファイル選択をキャンセルしました。|File selection canceled.|已取消文件选择。
ファイルを開けませんでした。|Could not open the file.|无法打开文件。
HTML ワークブックを開けませんでした。|Could not open the HTML workbook.|无法打开 HTML 工作簿。
HTML ワークブックを読み込みましたが、再読み込み用の保存に失敗しました。ファイルを保存してください。|Workbook loaded, but recovery storage failed. Please save the file.|工作簿已加载，但恢复保存失败。请保存文件。
`.trim().split('\n').map(line => { const [ja, en, zh] = line.split('|'); return [ja, { en, 'zh-CN': zh }]; }));
}
