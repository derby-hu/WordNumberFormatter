// 搜索模式：非数字非小数点字符 + 整数>=4位 + 小数点 + 恰好2位小数 + 非数字非小数点字符
// 修正：Word 通配符中 "." 表示任意字符，字面小数点必须写作 "\."
const SEARCH_PATTERN = "[!0-9.][0-9]{4,}\\.[0-9][0-9][!0-9.]";
const NUMBER_REGEX = /([0-9]{4,}\.[0-9][0-9])/;

/**
 * 为匹配到的文本添加千分符
 * @param {string} text 包含前后字符的匹配文本
 * @returns {string|null} 替换后的文本，不匹配则返回 null
 */
function formatMatchedText(text) {
    const match = text.match(NUMBER_REGEX);
    if (!match) return null;
    const numberText = match[1];
    const [intPart, decPart] = numberText.split(".");
    const formattedInt = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
    // 修正：原代码漏写 ${}，导致输出字面量 "formattedInt.{decPart}"
    return text.replace(numberText, `${formattedInt}.${decPart}`);
}

/**
 * 处理正文部分的数字（自动排除表格）
 * @returns {Promise<number>} 处理数量
 */
async function processBodyNumbers() {
    return await Word.run(async (context) => {
        const body = context.document.body;
        const results = body.search(SEARCH_PATTERN, { matchWildcards: true });
        results.load("text");
        await context.sync();
        // 判断每个搜索结果是否在表格内
        for (const range of results.items) {
            range.parentTableOrNullObject.load("isNullObject");
        }
        await context.sync();
        const replacements = [];
        for (const range of results.items) {
            if (!range.parentTableOrNullObject.isNullObject) continue;
            const newText = formatMatchedText(range.text);
            if (newText !== null) {
                replacements.push({ range, newText });
            }
        }
        // 批量提交替换
        for (const { range, newText } of replacements) {
            range.insertText(newText, Word.InsertLocation.replace);
        }
        await context.sync();
        return replacements.length;
    });
}

/**
 * 处理表格部分的数字
 * 修正：Word.Table 没有 cells 属性，必须经 table.rows -> row.cells 遍历单元格
 * @returns {Promise<number>} 处理数量
 */
async function processTableNumbers() {
    return await Word.run(async (context) => {
        const body = context.document.body;
        const tables = body.tables;
        tables.load("items");
        await context.sync();
        // 加载所有表格的行
        for (const table of tables.items) {
            table.rows.load("items");
        }
        await context.sync();
        // 加载所有行的单元格
        for (const table of tables.items) {
            for (const row of table.rows.items) {
                row.cells.load("items");
            }
        }
        await context.sync();
        // 收集所有单元格 body
        const cellBodies = [];
        for (const table of tables.items) {
            for (const row of table.rows.items) {
                for (const cell of row.cells.items) {
                    cellBodies.push(cell.body);
                }
            }
        }
        // 在每个单元格中搜索
        const searchResults = [];
        for (const cellBody of cellBodies) {
            const results = cellBody.search(SEARCH_PATTERN, { matchWildcards: true });
            results.load("text");
            searchResults.push(results);
        }
        await context.sync();
        const replacements = [];
        for (const results of searchResults) {
            for (const range of results.items) {
                const newText = formatMatchedText(range.text);
                if (newText !== null) {
                    replacements.push({ range, newText });
                }
            }
        }
        // 批量提交替换
        for (const { range, newText } of replacements) {
            range.insertText(newText, Word.InsertLocation.replace);
        }
        await context.sync();
        return replacements.length;
    });
}

/**
 * 一次性处理正文和表格中的所有数字
 * @returns {Promise<{bodyCount: number, tableCount: number}>}
 */
async function processAllNumbers() {
    const bodyCount = await processBodyNumbers();
    const tableCount = await processTableNumbers();
    return { bodyCount, tableCount };
}

// 修正：共享模块需经 webpack 打包，不能再用 <script src> 引入，故导出供 import 使用
export { processBodyNumbers, processTableNumbers, processAllNumbers };