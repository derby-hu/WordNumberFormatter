// 修正：通过 import 引入共享逻辑（webpack 打包），不再用 <script src>
import { processAllNumbers, processBodyNumbers, processTableNumbers } from "../shared/number-formatter.js";

Office.onReady((info) => {
    if (info.host !== Office.HostType.Word) return;

    const statusEl = document.getElementById("status");

    document.getElementById("format-all").onclick = async () => {
        statusEl.textContent = "正在处理全部数字...";
        try {
            const { bodyCount, tableCount } = await processAllNumbers();
            statusEl.textContent = `处理完成！正文 ${bodyCount} 处，表格 ${tableCount} 处，共 ${bodyCount + tableCount} 处。`;
        } catch (err) {
            statusEl.textContent = `处理失败：${err.message}`;
            console.error(err);
        }
    };

    document.getElementById("format-body").onclick = async () => {
        statusEl.textContent = "正在处理文档部分...";
        try {
            const count = await processBodyNumbers();
            statusEl.textContent = `文档部分处理完成！共 ${count} 处。`;
        } catch (err) {
            statusEl.textContent = `处理失败：${err.message}`;
            console.error(err);
        }
    };

    document.getElementById("format-table").onclick = async () => {
        statusEl.textContent = "正在处理表格部分...";
        try {
            const count = await processTableNumbers();
            statusEl.textContent = `表格部分处理完成！共 ${count} 处。`;
        } catch (err) {
            statusEl.textContent = `处理失败：${err.message}`;
            console.error(err);
        }
    };
});