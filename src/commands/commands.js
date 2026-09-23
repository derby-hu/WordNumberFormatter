// 修正：通过 import 引入共享逻辑（webpack 打包），不再用 <script src>
import { processAllNumbers, processBodyNumbers, processTableNumbers } from "../shared/number-formatter.js";

/**
 * 主按钮"千分符"：同时替换正文和表格中的数字
 */
function formatAllNumbers(event) {
    processAllNumbers()
        .then(({ bodyCount, tableCount }) => {
            // 修正：原代码日志漏写 ${}，导致输出字面量 bodyCount/tableCount
            console.log(`处理完成：正文 ${bodyCount} 处，表格 ${tableCount} 处，共 ${bodyCount + tableCount} 处`);
            event.completed();
        })
        .catch((err) => {
            console.error("处理失败：", err);
            event.completed();
        });
}

/**
 * 下拉菜单"文字替换"：仅处理正文
 */
function formatBodyNumbers(event) {
    processBodyNumbers()
        .then((count) => {
            console.log(`正文处理完成，共 ${count} 处`);
            event.completed();
        })
        .catch((err) => {
            console.error("正文处理失败：", err);
            event.completed();
        });
}

/**
 * 下拉菜单"表格替换"：仅处理表格
 */
function formatTableNumbers(event) {
    processTableNumbers()
        .then((count) => {
            console.log(`表格处理完成，共 ${count} 处`);
            event.completed();
        })
        .catch((err) => {
            console.error("表格处理失败：", err);
            event.completed();
        });
}

// 修正：ExecuteFunction 在 webpack 打包下函数不再是全局对象，
// 必须用 Office.actions.associate 将 actionId 与函数显式关联，
// actionId 需与 manifest 中的 <FunctionName> 完全一致
Office.actions.associate("formatAllNumbers", formatAllNumbers);
Office.actions.associate("formatBodyNumbers", formatBodyNumbers);
Office.actions.associate("formatTableNumbers", formatTableNumbers);