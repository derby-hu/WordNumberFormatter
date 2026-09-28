// Word 通配符特殊字符：[ ] ( ) { } * ? < > ! @
// "." 不是通配符，匹配字面小数点，无需转义。
// 之前的 BUG#2 误判 "." 为通配符，加 "\\." 反而导致搜索 "\." 两字符而匹配失败。

// 小数模式：边界 + 整数>=4位 + 小数点 + 恰好2位小数 + 边界
const SEARCH_PATTERN_DECIMAL = "[!0-9.][0-9]{4,}.[0-9][0-9][!0-9.]";
// 纯整数模式：边界 + 整数>=4位 + 边界（不跟小数点，因 [!0-9.] 排除了 "."）
const SEARCH_PATTERN_INTEGER = "[!0-9.][0-9]{4,}[!0-9.]";

const NUMBER_REGEX_DECIMAL = /([0-9]{4,}\.[0-9][0-9])/;
const NUMBER_REGEX_INTEGER = /([0-9]{4,})/;

/**
 * 为匹配到的文本添加千分符
 * @param {string} text 包含前后字符的匹配文本
 * @returns {string|null} 替换后的文本，不匹配则返回 null
 */
function formatMatchedText(text) {
    // 先尝试匹配小数
    let match = text.match(NUMBER_REGEX_DECIMAL);
    if (match) {
        const numberText = match[1];
        const [intPart, decPart] = numberText.split(".");
        const formattedInt = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
        return text.replace(numberText, `${formattedInt}.${decPart}`);
    }
    // 再尝试匹配纯整数
    match = text.match(NUMBER_REGEX_INTEGER);
    if (match) {
        const numberText = match[1];
        const formattedInt = numberText.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
        return text.replace(numberText, formattedInt);
    }
    return null;
}

/**
 * 在指定搜索范围内执行两种模式的搜索并返回所有结果范围
 * @param {Word.Range} searchScope body 或 cell.body
 * @returns {Promise<Word.Range[]>} 匹配的范围数组
 */
async function searchNumbers(context, searchScope) {
    const decimalResults = searchScope.search(SEARCH_PATTERN_DECIMAL, { matchWildcards: true });
    const integerResults = searchScope.search(SEARCH_PATTERN_INTEGER, { matchWildcards: true });
    decimalResults.load("text");
    integerResults.load("text");
    await context.sync();
    return [...decimalResults.items, ...integerResults.items];
}

/**
 * 处理正文部分的数字（自动排除表格）
 * @returns {Promise<number>} 处理数量
 */
async function processBodyNumbers() {
    return await Word.run(async (context) => {
        const body = context.document.body;
        const ranges = await searchNumbers(context, body);
        // 判断每个搜索结果是否在表格内
        for (const range of ranges) {
            range.parentTableOrNullObject.load("isNullObject");
        }
        await context.sync();
        const replacements = [];
        for (const range of ranges) {
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
 * 处理表格部分的数字（含嵌套表格）
 * Word.Table 没有 cells 属性，必须经 table.rows -> row.cells 遍历单元格；
 * body.tables 仅含顶层表格，嵌套表格通过 table.tables 逐层递归处理
 * @returns {Promise<number>} 处理数量
 */
async function processTableNumbers() {
    return await Word.run(async (context) => {
        const body = context.document.body;
        const tables = body.tables;
        tables.load("items");
        await context.sync();
        return await processTableLevel(context, tables.items);
    });
}

/**
 * 递归处理一层表格：先搜索替换本层所有单元格并提交，再深入嵌套表格。
 * 逐层提交可避免外层与内层的匹配范围重叠导致重复替换；
 * 已格式化的数字因含逗号不再满足搜索模式，天然幂等。
 * @param {Word.RequestContext} context
 * @param {Word.Table[]} tables 当前层的表格集合
 * @returns {Promise<number>} 本层及所有嵌套层累计处理数量
 */
async function processTableLevel(context, tables) {
    // 加载本层所有表格的行与嵌套表格集合
    for (const table of tables) {
        table.rows.load("items");
        table.tables.load("items");
    }
    await context.sync();
    for (const table of tables) {
        for (const row of table.rows.items) {
            row.cells.load("items");
        }
    }
    await context.sync();

    // 在本层每个单元格中搜索两种模式
    const allRanges = [];
    for (const table of tables) {
        for (const row of table.rows.items) {
            for (const cell of row.cells.items) {
                const ranges = await searchNumbers(context, cell.body);
                allRanges.push(...ranges);
            }
        }
    }

    // 批量提交替换
    const replacements = [];
    for (const range of allRanges) {
        const newText = formatMatchedText(range.text);
        if (newText !== null) {
            replacements.push({ range, newText });
        }
    }
    for (const { range, newText } of replacements) {
        range.insertText(newText, Word.InsertLocation.replace);
    }
    await context.sync();
    let count = replacements.length;

    // 递归处理嵌套表格
    for (const table of tables) {
        if (table.tables.items.length > 0) {
            count += await processTableLevel(context, table.tables.items);
        }
    }
    return count;
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

export { processBodyNumbers, processTableNumbers, processAllNumbers };
