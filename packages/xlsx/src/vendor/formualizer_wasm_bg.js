export class ASTNode {
    static __wrap(ptr) {
        ptr = ptr >>> 0;
        const obj = Object.create(ASTNode.prototype);
        obj.__wbg_ptr = ptr;
        ASTNodeFinalization.register(obj, obj.__wbg_ptr, obj);
        return obj;
    }
    __destroy_into_raw() {
        const ptr = this.__wbg_ptr;
        this.__wbg_ptr = 0;
        ASTNodeFinalization.unregister(this);
        return ptr;
    }
    free() {
        const ptr = this.__destroy_into_raw();
        wasm.__wbg_astnode_free(ptr, 0);
    }
    /**
     * @returns {string}
     */
    getType() {
        let deferred1_0;
        let deferred1_1;
        try {
            const ret = wasm.astnode_getType(this.__wbg_ptr);
            deferred1_0 = ret[0];
            deferred1_1 = ret[1];
            return getStringFromWasm0(ret[0], ret[1]);
        } finally {
            wasm.__wbindgen_free(deferred1_0, deferred1_1, 1);
        }
    }
    /**
     * @returns {any}
     */
    toJSON() {
        const ret = wasm.astnode_toJSON(this.__wbg_ptr);
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        return takeFromExternrefTable0(ret[0]);
    }
    /**
     * @returns {string}
     */
    toString() {
        let deferred1_0;
        let deferred1_1;
        try {
            const ret = wasm.astnode_toString(this.__wbg_ptr);
            deferred1_0 = ret[0];
            deferred1_1 = ret[1];
            return getStringFromWasm0(ret[0], ret[1]);
        } finally {
            wasm.__wbindgen_free(deferred1_0, deferred1_1, 1);
        }
    }
}
if (Symbol.dispose) ASTNode.prototype[Symbol.dispose] = ASTNode.prototype.free;

/**
 * @enum {0 | 1}
 */
export const FormulaDialect = Object.freeze({
    Excel: 0, "0": "Excel",
    OpenFormula: 1, "1": "OpenFormula",
});

export class JsToken {
    __destroy_into_raw() {
        const ptr = this.__wbg_ptr;
        this.__wbg_ptr = 0;
        JsTokenFinalization.unregister(this);
        return ptr;
    }
    free() {
        const ptr = this.__destroy_into_raw();
        wasm.__wbg_jstoken_free(ptr, 0);
    }
    /**
     * @returns {number}
     */
    get end() {
        const ret = wasm.jstoken_end(this.__wbg_ptr);
        return ret >>> 0;
    }
    /**
     * @returns {number}
     */
    get pos() {
        const ret = wasm.jstoken_pos(this.__wbg_ptr);
        return ret >>> 0;
    }
    /**
     * @returns {string}
     */
    get subtype() {
        let deferred1_0;
        let deferred1_1;
        try {
            const ret = wasm.jstoken_subtype(this.__wbg_ptr);
            deferred1_0 = ret[0];
            deferred1_1 = ret[1];
            return getStringFromWasm0(ret[0], ret[1]);
        } finally {
            wasm.__wbindgen_free(deferred1_0, deferred1_1, 1);
        }
    }
    /**
     * @returns {any}
     */
    toJSON() {
        const ret = wasm.jstoken_toJSON(this.__wbg_ptr);
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        return takeFromExternrefTable0(ret[0]);
    }
    /**
     * @returns {string}
     */
    toString() {
        let deferred1_0;
        let deferred1_1;
        try {
            const ret = wasm.jstoken_toString(this.__wbg_ptr);
            deferred1_0 = ret[0];
            deferred1_1 = ret[1];
            return getStringFromWasm0(ret[0], ret[1]);
        } finally {
            wasm.__wbindgen_free(deferred1_0, deferred1_1, 1);
        }
    }
    /**
     * @returns {string}
     */
    get tokenType() {
        let deferred1_0;
        let deferred1_1;
        try {
            const ret = wasm.jstoken_tokenType(this.__wbg_ptr);
            deferred1_0 = ret[0];
            deferred1_1 = ret[1];
            return getStringFromWasm0(ret[0], ret[1]);
        } finally {
            wasm.__wbindgen_free(deferred1_0, deferred1_1, 1);
        }
    }
    /**
     * @returns {string}
     */
    get value() {
        let deferred1_0;
        let deferred1_1;
        try {
            const ret = wasm.jstoken_value(this.__wbg_ptr);
            deferred1_0 = ret[0];
            deferred1_1 = ret[1];
            return getStringFromWasm0(ret[0], ret[1]);
        } finally {
            wasm.__wbindgen_free(deferred1_0, deferred1_1, 1);
        }
    }
}
if (Symbol.dispose) JsToken.prototype[Symbol.dispose] = JsToken.prototype.free;

export class Parser {
    __destroy_into_raw() {
        const ptr = this.__wbg_ptr;
        this.__wbg_ptr = 0;
        ParserFinalization.unregister(this);
        return ptr;
    }
    free() {
        const ptr = this.__destroy_into_raw();
        wasm.__wbg_parser_free(ptr, 0);
    }
    /**
     * @param {string} formula
     * @param {FormulaDialect | null} [dialect]
     */
    constructor(formula, dialect) {
        const ptr0 = passStringToWasm0(formula, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.parser_new(ptr0, len0, isLikeNone(dialect) ? 2 : dialect);
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        this.__wbg_ptr = ret[0] >>> 0;
        ParserFinalization.register(this, this.__wbg_ptr, this);
        return this;
    }
    /**
     * @returns {ASTNode}
     */
    parse() {
        const ret = wasm.parser_parse(this.__wbg_ptr);
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        return ASTNode.__wrap(ret[0]);
    }
}
if (Symbol.dispose) Parser.prototype[Symbol.dispose] = Parser.prototype.free;

export class ParserError {
    __destroy_into_raw() {
        const ptr = this.__wbg_ptr;
        this.__wbg_ptr = 0;
        ParserErrorFinalization.unregister(this);
        return ptr;
    }
    free() {
        const ptr = this.__destroy_into_raw();
        wasm.__wbg_parsererror_free(ptr, 0);
    }
    /**
     * @returns {string}
     */
    get message() {
        let deferred1_0;
        let deferred1_1;
        try {
            const ret = wasm.parsererror_message(this.__wbg_ptr);
            deferred1_0 = ret[0];
            deferred1_1 = ret[1];
            return getStringFromWasm0(ret[0], ret[1]);
        } finally {
            wasm.__wbindgen_free(deferred1_0, deferred1_1, 1);
        }
    }
    /**
     * @param {string} message
     * @param {number | null} [pos]
     */
    constructor(message, pos) {
        const ptr0 = passStringToWasm0(message, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.parsererror_new(ptr0, len0, isLikeNone(pos) ? 0x100000001 : (pos) >>> 0);
        this.__wbg_ptr = ret >>> 0;
        ParserErrorFinalization.register(this, this.__wbg_ptr, this);
        return this;
    }
    /**
     * @returns {number | undefined}
     */
    get pos() {
        const ret = wasm.parsererror_pos(this.__wbg_ptr);
        return ret === 0x100000001 ? undefined : ret;
    }
    /**
     * @returns {string}
     */
    toString() {
        let deferred1_0;
        let deferred1_1;
        try {
            const ret = wasm.parsererror_toString(this.__wbg_ptr);
            deferred1_0 = ret[0];
            deferred1_1 = ret[1];
            return getStringFromWasm0(ret[0], ret[1]);
        } finally {
            wasm.__wbindgen_free(deferred1_0, deferred1_1, 1);
        }
    }
}
if (Symbol.dispose) ParserError.prototype[Symbol.dispose] = ParserError.prototype.free;

export class Reference {
    __destroy_into_raw() {
        const ptr = this.__wbg_ptr;
        this.__wbg_ptr = 0;
        ReferenceFinalization.unregister(this);
        return ptr;
    }
    free() {
        const ptr = this.__destroy_into_raw();
        wasm.__wbg_reference_free(ptr, 0);
    }
    /**
     * @returns {boolean}
     */
    get colAbsEnd() {
        const ret = wasm.reference_colAbsEnd(this.__wbg_ptr);
        return ret !== 0;
    }
    /**
     * @returns {boolean}
     */
    get colAbsStart() {
        const ret = wasm.reference_colAbsStart(this.__wbg_ptr);
        return ret !== 0;
    }
    /**
     * @returns {number}
     */
    get colEnd() {
        const ret = wasm.reference_colEnd(this.__wbg_ptr);
        return ret >>> 0;
    }
    /**
     * @returns {number}
     */
    get colStart() {
        const ret = wasm.reference_colStart(this.__wbg_ptr);
        return ret >>> 0;
    }
    /**
     * @returns {boolean}
     */
    isRange() {
        const ret = wasm.reference_isRange(this.__wbg_ptr);
        return ret !== 0;
    }
    /**
     * @returns {boolean}
     */
    isSingleCell() {
        const ret = wasm.reference_isSingleCell(this.__wbg_ptr);
        return ret !== 0;
    }
    /**
     * @param {string | null | undefined} sheet
     * @param {number} row_start
     * @param {number} col_start
     * @param {number} row_end
     * @param {number} col_end
     * @param {boolean} row_abs_start
     * @param {boolean} col_abs_start
     * @param {boolean} row_abs_end
     * @param {boolean} col_abs_end
     */
    constructor(sheet, row_start, col_start, row_end, col_end, row_abs_start, col_abs_start, row_abs_end, col_abs_end) {
        var ptr0 = isLikeNone(sheet) ? 0 : passStringToWasm0(sheet, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        var len0 = WASM_VECTOR_LEN;
        const ret = wasm.reference_new(ptr0, len0, row_start, col_start, row_end, col_end, row_abs_start, col_abs_start, row_abs_end, col_abs_end);
        this.__wbg_ptr = ret >>> 0;
        ReferenceFinalization.register(this, this.__wbg_ptr, this);
        return this;
    }
    /**
     * @returns {boolean}
     */
    get rowAbsEnd() {
        const ret = wasm.reference_rowAbsEnd(this.__wbg_ptr);
        return ret !== 0;
    }
    /**
     * @returns {boolean}
     */
    get rowAbsStart() {
        const ret = wasm.reference_rowAbsStart(this.__wbg_ptr);
        return ret !== 0;
    }
    /**
     * @returns {number}
     */
    get rowEnd() {
        const ret = wasm.reference_rowEnd(this.__wbg_ptr);
        return ret >>> 0;
    }
    /**
     * @returns {number}
     */
    get rowStart() {
        const ret = wasm.reference_rowStart(this.__wbg_ptr);
        return ret >>> 0;
    }
    /**
     * @returns {string | undefined}
     */
    get sheet() {
        const ret = wasm.reference_sheet(this.__wbg_ptr);
        let v1;
        if (ret[0] !== 0) {
            v1 = getStringFromWasm0(ret[0], ret[1]).slice();
            wasm.__wbindgen_free(ret[0], ret[1] * 1, 1);
        }
        return v1;
    }
    /**
     * @returns {any}
     */
    toJSON() {
        const ret = wasm.reference_toJSON(this.__wbg_ptr);
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        return takeFromExternrefTable0(ret[0]);
    }
    /**
     * @returns {string}
     */
    toString() {
        let deferred1_0;
        let deferred1_1;
        try {
            const ret = wasm.reference_toString(this.__wbg_ptr);
            deferred1_0 = ret[0];
            deferred1_1 = ret[1];
            return getStringFromWasm0(ret[0], ret[1]);
        } finally {
            wasm.__wbindgen_free(deferred1_0, deferred1_1, 1);
        }
    }
}
if (Symbol.dispose) Reference.prototype[Symbol.dispose] = Reference.prototype.free;

export class Sheet {
    static __wrap(ptr) {
        ptr = ptr >>> 0;
        const obj = Object.create(Sheet.prototype);
        obj.__wbg_ptr = ptr;
        SheetFinalization.register(obj, obj.__wbg_ptr, obj);
        return obj;
    }
    __destroy_into_raw() {
        const ptr = this.__wbg_ptr;
        this.__wbg_ptr = 0;
        SheetFinalization.unregister(this);
        return ptr;
    }
    free() {
        const ptr = this.__destroy_into_raw();
        wasm.__wbg_sheet_free(ptr, 0);
    }
    /**
     * @param {number} row
     * @param {number} col
     * @returns {any}
     */
    evaluateCell(row, col) {
        const ret = wasm.sheet_evaluateCell(this.__wbg_ptr, row, col);
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        return takeFromExternrefTable0(ret[0]);
    }
    /**
     * @param {number} row
     * @param {number} col
     * @returns {string | undefined}
     */
    getFormula(row, col) {
        const ret = wasm.sheet_getFormula(this.__wbg_ptr, row, col);
        if (ret[3]) {
            throw takeFromExternrefTable0(ret[2]);
        }
        let v1;
        if (ret[0] !== 0) {
            v1 = getStringFromWasm0(ret[0], ret[1]).slice();
            wasm.__wbindgen_free(ret[0], ret[1] * 1, 1);
        }
        return v1;
    }
    /**
     * @param {number} row
     * @param {number} col
     * @returns {any}
     */
    getValue(row, col) {
        const ret = wasm.sheet_getValue(this.__wbg_ptr, row, col);
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        return takeFromExternrefTable0(ret[0]);
    }
    /**
     * Read a rectangular range of values as a 2D array (no evaluation)
     * @param {number} start_row
     * @param {number} start_col
     * @param {number} end_row
     * @param {number} end_col
     * @returns {Array<any>}
     */
    readRange(start_row, start_col, end_row, end_col) {
        const ret = wasm.sheet_readRange(this.__wbg_ptr, start_row, start_col, end_row, end_col);
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        return takeFromExternrefTable0(ret[0]);
    }
    /**
     * @param {number} row
     * @param {number} col
     * @param {string} formula
     */
    setFormula(row, col, formula) {
        const ptr0 = passStringToWasm0(formula, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.sheet_setFormula(this.__wbg_ptr, row, col, ptr0, len0);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * @param {number} start_row
     * @param {number} start_col
     * @param {Array<any>} data
     */
    setFormulas(start_row, start_col, data) {
        const ret = wasm.sheet_setFormulas(this.__wbg_ptr, start_row, start_col, data);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * @param {number} row
     * @param {number} col
     * @param {any} value
     */
    setValue(row, col, value) {
        const ret = wasm.sheet_setValue(this.__wbg_ptr, row, col, value);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * @param {number} start_row
     * @param {number} start_col
     * @param {Array<any>} data
     */
    setValues(start_row, start_col, data) {
        const ret = wasm.sheet_setValues(this.__wbg_ptr, start_row, start_col, data);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
}
if (Symbol.dispose) Sheet.prototype[Symbol.dispose] = Sheet.prototype.free;

export class SheetPortSession {
    static __wrap(ptr) {
        ptr = ptr >>> 0;
        const obj = Object.create(SheetPortSession.prototype);
        obj.__wbg_ptr = ptr;
        SheetPortSessionFinalization.register(obj, obj.__wbg_ptr, obj);
        return obj;
    }
    __destroy_into_raw() {
        const ptr = this.__wbg_ptr;
        this.__wbg_ptr = 0;
        SheetPortSessionFinalization.unregister(this);
        return ptr;
    }
    free() {
        const ptr = this.__destroy_into_raw();
        wasm.__wbg_sheetportsession_free(ptr, 0);
    }
    /**
     * @returns {any}
     */
    describePorts() {
        const ret = wasm.sheetportsession_describePorts(this.__wbg_ptr);
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        return takeFromExternrefTable0(ret[0]);
    }
    /**
     * @param {any} options
     * @returns {any}
     */
    evaluateOnce(options) {
        const ret = wasm.sheetportsession_evaluateOnce(this.__wbg_ptr, options);
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        return takeFromExternrefTable0(ret[0]);
    }
    /**
     * @param {any} value
     * @param {Workbook} workbook
     * @returns {SheetPortSession}
     */
    static fromManifest(value, workbook) {
        _assertClass(workbook, Workbook);
        const ret = wasm.sheetportsession_fromManifest(value, workbook.__wbg_ptr);
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        return SheetPortSession.__wrap(ret[0]);
    }
    /**
     * @param {string} manifest_yaml
     * @param {Workbook} workbook
     * @returns {SheetPortSession}
     */
    static fromManifestYaml(manifest_yaml, workbook) {
        const ptr0 = passStringToWasm0(manifest_yaml, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        _assertClass(workbook, Workbook);
        const ret = wasm.sheetportsession_fromManifestYaml(ptr0, len0, workbook.__wbg_ptr);
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        return SheetPortSession.__wrap(ret[0]);
    }
    /**
     * @returns {any}
     */
    manifest() {
        const ret = wasm.sheetportsession_manifest(this.__wbg_ptr);
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        return takeFromExternrefTable0(ret[0]);
    }
    /**
     * @returns {any}
     */
    readInputs() {
        const ret = wasm.sheetportsession_readInputs(this.__wbg_ptr);
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        return takeFromExternrefTable0(ret[0]);
    }
    /**
     * @returns {any}
     */
    readOutputs() {
        const ret = wasm.sheetportsession_readOutputs(this.__wbg_ptr);
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        return takeFromExternrefTable0(ret[0]);
    }
    /**
     * @param {any} updates
     */
    writeInputs(updates) {
        const ret = wasm.sheetportsession_writeInputs(this.__wbg_ptr, updates);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
}
if (Symbol.dispose) SheetPortSession.prototype[Symbol.dispose] = SheetPortSession.prototype.free;

export class Tokenizer {
    static __wrap(ptr) {
        ptr = ptr >>> 0;
        const obj = Object.create(Tokenizer.prototype);
        obj.__wbg_ptr = ptr;
        TokenizerFinalization.register(obj, obj.__wbg_ptr, obj);
        return obj;
    }
    __destroy_into_raw() {
        const ptr = this.__wbg_ptr;
        this.__wbg_ptr = 0;
        TokenizerFinalization.unregister(this);
        return ptr;
    }
    free() {
        const ptr = this.__destroy_into_raw();
        wasm.__wbg_tokenizer_free(ptr, 0);
    }
    /**
     * @param {number} index
     * @returns {any}
     */
    getToken(index) {
        const ret = wasm.tokenizer_getToken(this.__wbg_ptr, index);
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        return takeFromExternrefTable0(ret[0]);
    }
    /**
     * @returns {number}
     */
    length() {
        const ret = wasm.tokenizer_length(this.__wbg_ptr);
        return ret >>> 0;
    }
    /**
     * @param {string} formula
     * @param {FormulaDialect | null} [dialect]
     */
    constructor(formula, dialect) {
        const ptr0 = passStringToWasm0(formula, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.tokenizer_new(ptr0, len0, isLikeNone(dialect) ? 2 : dialect);
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        this.__wbg_ptr = ret[0] >>> 0;
        TokenizerFinalization.register(this, this.__wbg_ptr, this);
        return this;
    }
    /**
     * @returns {string}
     */
    render() {
        let deferred1_0;
        let deferred1_1;
        try {
            const ret = wasm.tokenizer_render(this.__wbg_ptr);
            deferred1_0 = ret[0];
            deferred1_1 = ret[1];
            return getStringFromWasm0(ret[0], ret[1]);
        } finally {
            wasm.__wbindgen_free(deferred1_0, deferred1_1, 1);
        }
    }
    /**
     * @returns {string}
     */
    toString() {
        let deferred1_0;
        let deferred1_1;
        try {
            const ret = wasm.tokenizer_toString(this.__wbg_ptr);
            deferred1_0 = ret[0];
            deferred1_1 = ret[1];
            return getStringFromWasm0(ret[0], ret[1]);
        } finally {
            wasm.__wbindgen_free(deferred1_0, deferred1_1, 1);
        }
    }
    /**
     * @returns {any}
     */
    tokens() {
        const ret = wasm.tokenizer_tokens(this.__wbg_ptr);
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        return takeFromExternrefTable0(ret[0]);
    }
}
if (Symbol.dispose) Tokenizer.prototype[Symbol.dispose] = Tokenizer.prototype.free;

export class TokenizerError {
    __destroy_into_raw() {
        const ptr = this.__wbg_ptr;
        this.__wbg_ptr = 0;
        TokenizerErrorFinalization.unregister(this);
        return ptr;
    }
    free() {
        const ptr = this.__destroy_into_raw();
        wasm.__wbg_tokenizererror_free(ptr, 0);
    }
    /**
     * @returns {string}
     */
    get message() {
        let deferred1_0;
        let deferred1_1;
        try {
            const ret = wasm.tokenizererror_message(this.__wbg_ptr);
            deferred1_0 = ret[0];
            deferred1_1 = ret[1];
            return getStringFromWasm0(ret[0], ret[1]);
        } finally {
            wasm.__wbindgen_free(deferred1_0, deferred1_1, 1);
        }
    }
    /**
     * @param {string} message
     * @param {number | null} [pos]
     */
    constructor(message, pos) {
        const ptr0 = passStringToWasm0(message, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.parsererror_new(ptr0, len0, isLikeNone(pos) ? 0x100000001 : (pos) >>> 0);
        this.__wbg_ptr = ret >>> 0;
        TokenizerErrorFinalization.register(this, this.__wbg_ptr, this);
        return this;
    }
    /**
     * @returns {number | undefined}
     */
    get pos() {
        const ret = wasm.parsererror_pos(this.__wbg_ptr);
        return ret === 0x100000001 ? undefined : ret;
    }
    /**
     * @returns {string}
     */
    toString() {
        let deferred1_0;
        let deferred1_1;
        try {
            const ret = wasm.tokenizererror_toString(this.__wbg_ptr);
            deferred1_0 = ret[0];
            deferred1_1 = ret[1];
            return getStringFromWasm0(ret[0], ret[1]);
        } finally {
            wasm.__wbindgen_free(deferred1_0, deferred1_1, 1);
        }
    }
}
if (Symbol.dispose) TokenizerError.prototype[Symbol.dispose] = TokenizerError.prototype.free;

export class Workbook {
    static __wrap(ptr) {
        ptr = ptr >>> 0;
        const obj = Object.create(Workbook.prototype);
        obj.__wbg_ptr = ptr;
        WorkbookFinalization.register(obj, obj.__wbg_ptr, obj);
        return obj;
    }
    __destroy_into_raw() {
        const ptr = this.__wbg_ptr;
        this.__wbg_ptr = 0;
        WorkbookFinalization.unregister(this);
        return ptr;
    }
    free() {
        const ptr = this.__destroy_into_raw();
        wasm.__wbg_workbook_free(ptr, 0);
    }
    /**
     * @param {string} name
     */
    addSheet(name) {
        const ptr0 = passStringToWasm0(name, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.workbook_addSheet(this.__wbg_ptr, ptr0, len0);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * Define a native table over cells that already exist.
     *
     * `definition` is `{ name, sheet, range: [firstRow, firstCol, lastRow,
     * lastCol], headers: string[], headerRow?: boolean, totalsRow?: boolean }`.
     * `range` is 1-based and inclusive, and covers the header row when
     * `headerRow` is true (default `true`); `totalsRow` defaults to `false`.
     *
     * Tables are metadata over existing cells, so populate the region first.
     * Unknown keys are rejected rather than ignored.
     * @param {TableDefinition} definition
     */
    addTable(definition) {
        const ret = wasm.workbook_addTable(this.__wbg_ptr, definition);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * @param {string} description
     */
    beginAction(description) {
        const ptr0 = passStringToWasm0(description, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.workbook_beginAction(this.__wbg_ptr, ptr0, len0);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    cancel() {
        wasm.workbook_cancel(this.__wbg_ptr);
    }
    /**
     * Return bounded dependents in canonical address order.
     * @param {any} cell
     * @param {any | null} [options]
     * @returns {any}
     */
    dependents(cell, options) {
        const ret = wasm.workbook_dependents(this.__wbg_ptr, cell, isLikeNone(options) ? 0 : addToExternrefTable0(options));
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        return takeFromExternrefTable0(ret[0]);
    }
    endAction() {
        const ret = wasm.workbook_endAction(this.__wbg_ptr);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    evaluateAll() {
        const ret = wasm.workbook_evaluateAll(this.__wbg_ptr);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * @param {string} sheet
     * @param {number} row
     * @param {number} col
     * @returns {any}
     */
    evaluateCell(sheet, row, col) {
        const ptr0 = passStringToWasm0(sheet, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.workbook_evaluateCell(this.__wbg_ptr, ptr0, len0, row, col);
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        return takeFromExternrefTable0(ret[0]);
    }
    /**
     * @param {Array<any>} targets
     * @returns {Array<any>}
     */
    evaluateCells(targets) {
        const ret = wasm.workbook_evaluateCells(this.__wbg_ptr, targets);
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        return takeFromExternrefTable0(ret[0]);
    }
    /**
     * Construct from a JSON workbook string (feature: json)
     * @param {string} json
     * @returns {Workbook}
     */
    static fromJson(json) {
        const ptr0 = passStringToWasm0(json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.workbook_fromJson(ptr0, len0);
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        return Workbook.__wrap(ret[0]);
    }
    /**
     * Construct from a JSON workbook string with options, e.g.
     * `{ spanEvaluation: true }` to opt into experimental FormulaPlane spans.
     * @param {string} json
     * @param {any | null} [options]
     * @returns {Workbook}
     */
    static fromJsonWithOptions(json, options) {
        const ptr0 = passStringToWasm0(json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.workbook_fromJsonWithOptions(ptr0, len0, isLikeNone(options) ? 0 : addToExternrefTable0(options));
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        return Workbook.__wrap(ret[0]);
    }
    /**
     * Construct from XLSX bytes using the native Compatible reader.
     * @param {Uint8Array} bytes
     * @returns {Workbook}
     */
    static fromXlsxBytes(bytes) {
        const ptr0 = passArray8ToWasm0(bytes, wasm.__wbindgen_malloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.workbook_fromXlsxBytes(ptr0, len0);
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        return Workbook.__wrap(ret[0]);
    }
    /**
     * Construct from XLSX bytes with options, e.g. `{ backend: "native" }`
     * for the native reader, or `{ spanEvaluation: true }` for spans.
     * Native is the default; `calamine` is a compatibility alias.
     * `loadLimits` accepts optional maxInputBytes, maxEntries, maxExpandedBytes,
     * maxPartBytes and maxXmlDepth non-negative safe integers within host width.
     * Partial limits inherit Rust defaults; zero means zero. Unknown keys fail.
     * Limits apply before core admission, after the JS-to-WASM byte transfer.
     * @param {Uint8Array} bytes
     * @param {any | null} [options]
     * @returns {Workbook}
     */
    static fromXlsxBytesWithOptions(bytes, options) {
        const ptr0 = passArray8ToWasm0(bytes, wasm.__wbindgen_malloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.workbook_fromXlsxBytesWithOptions(ptr0, len0, isLikeNone(options) ? 0 : addToExternrefTable0(options));
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        return Workbook.__wrap(ret[0]);
    }
    /**
     * @param {Array<any>} targets
     * @param {any | null} [options]
     * @returns {any}
     */
    getEvalPlan(targets, options) {
        const ret = wasm.workbook_getEvalPlan(this.__wbg_ptr, targets, isLikeNone(options) ? 0 : addToExternrefTable0(options));
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        return takeFromExternrefTable0(ret[0]);
    }
    /**
     * @param {string | null} [sheet]
     * @returns {Array<any>}
     */
    getNamedRanges(sheet) {
        var ptr0 = isLikeNone(sheet) ? 0 : passStringToWasm0(sheet, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        var len0 = WASM_VECTOR_LEN;
        const ret = wasm.workbook_getNamedRanges(this.__wbg_ptr, ptr0, len0);
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        return takeFromExternrefTable0(ret[0]);
    }
    /**
     * Metadata for every defined table, ordered by name.
     * @returns {TableMetadata[]}
     */
    getTables() {
        const ret = wasm.workbook_getTables(this.__wbg_ptr);
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        return takeFromExternrefTable0(ret[0]);
    }
    /**
     * Return an owned, read-only snapshot of one cell.
     * @param {any} cell
     * @param {any | null} [options]
     * @returns {any}
     */
    inspectCell(cell, options) {
        const ret = wasm.workbook_inspectCell(this.__wbg_ptr, cell, isLikeNone(options) ? 0 : addToExternrefTable0(options));
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        return takeFromExternrefTable0(ret[0]);
    }
    /**
     * Telemetry from runtime SCC / iterative-calculation evaluation during
     * the most recent evaluation request (RFC #113, spec §10).
     *
     * Mirrors the engine accessor of the same name: counters reset at the
     * start of every evaluation request, so this always describes the LAST
     * `evaluateAll()` / `evaluateCell(s)` call. All-zero when cycle
     * detection is `"static"` or nothing cyclic was evaluated.
     * @returns {any}
     */
    lastCycleTelemetry() {
        const ret = wasm.workbook_lastCycleTelemetry(this.__wbg_ptr);
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        return takeFromExternrefTable0(ret[0]);
    }
    /**
     * List workbook-local custom functions and metadata.
     * @returns {Array<any>}
     */
    listFunctions() {
        const ret = wasm.workbook_listFunctions(this.__wbg_ptr);
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        return takeFromExternrefTable0(ret[0]);
    }
    /**
     * Construct an empty workbook, optionally with load options, e.g.
     * `{ cyclePolicy: "iterate", iterateMaxIterations: 50 }` to enable
     * iterative calculation for circular references (RFC #113, spec §2).
     *
     * `new Workbook()` (no arguments) behaves exactly as before: the
     * missing/undefined/null options map to the default interactive config.
     * @param {any | null} [options]
     */
    constructor(options) {
        const ret = wasm.workbook_new(isLikeNone(options) ? 0 : addToExternrefTable0(options));
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        this.__wbg_ptr = ret[0] >>> 0;
        WorkbookFinalization.register(this, this.__wbg_ptr, this);
        return this;
    }
    /**
     * Return source-ordered, first-occurrence-deduplicated declared precedents.
     * @param {any} cell
     * @param {any | null} [options]
     * @returns {any}
     */
    precedents(cell, options) {
        const ret = wasm.workbook_precedents(this.__wbg_ptr, cell, isLikeNone(options) ? 0 : addToExternrefTable0(options));
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        return takeFromExternrefTable0(ret[0]);
    }
    /**
     * Return a row-major page over a finite or semantically resolved open area.
     * @param {any} area
     * @param {any | null} [options]
     * @returns {any}
     */
    rangePage(area, options) {
        const ret = wasm.workbook_rangePage(this.__wbg_ptr, area, isLikeNone(options) ? 0 : addToExternrefTable0(options));
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        return takeFromExternrefTable0(ret[0]);
    }
    redo() {
        const ret = wasm.workbook_redo(this.__wbg_ptr);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * Register a workbook-local custom function backed by a JavaScript callback.
     * @param {string} name
     * @param {Function} callback
     * @param {any | null} [options]
     */
    registerFunction(name, callback, options) {
        const ptr0 = passStringToWasm0(name, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.workbook_registerFunction(this.__wbg_ptr, ptr0, len0, callback, isLikeNone(options) ? 0 : addToExternrefTable0(options));
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    resetCancel() {
        wasm.workbook_resetCancel(this.__wbg_ptr);
    }
    /**
     * @param {boolean} enabled
     */
    setChangelogEnabled(enabled) {
        const ret = wasm.workbook_setChangelogEnabled(this.__wbg_ptr, enabled);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * @param {string} sheet
     * @param {number} row
     * @param {number} col
     * @param {string} formula
     */
    setFormula(sheet, row, col, formula) {
        const ptr0 = passStringToWasm0(sheet, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ptr1 = passStringToWasm0(formula, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len1 = WASM_VECTOR_LEN;
        const ret = wasm.workbook_setFormula(this.__wbg_ptr, ptr0, len0, row, col, ptr1, len1);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * Choose temporal output as native JS dates (default) or numeric serials.
     * @param {string} policy
     */
    setTemporalEgress(policy) {
        const ptr0 = passStringToWasm0(policy, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.workbook_setTemporalEgress(this.__wbg_ptr, ptr0, len0);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * @param {string} sheet
     * @param {number} row
     * @param {number} col
     * @param {any} value
     */
    setValue(sheet, row, col, value) {
        const ptr0 = passStringToWasm0(sheet, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.workbook_setValue(this.__wbg_ptr, ptr0, len0, row, col, value);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * Get a sheet facade by name (creates if missing)
     * @param {string} name
     * @returns {Sheet}
     */
    sheet(name) {
        const ptr0 = passStringToWasm0(name, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.workbook_sheet(this.__wbg_ptr, ptr0, len0);
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        return Sheet.__wrap(ret[0]);
    }
    /**
     * @returns {Array<any>}
     */
    sheetNames() {
        const ret = wasm.workbook_sheetNames(this.__wbg_ptr);
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        return takeFromExternrefTable0(ret[0]);
    }
    /**
     * Build a bounded, owned trace graph for one or more roots.
     * @param {any} roots
     * @param {any | null} [options]
     * @returns {any}
     */
    trace(roots, options) {
        const ret = wasm.workbook_trace(this.__wbg_ptr, roots, isLikeNone(options) ? 0 : addToExternrefTable0(options));
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        return takeFromExternrefTable0(ret[0]);
    }
    undo() {
        const ret = wasm.workbook_undo(this.__wbg_ptr);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * Unregister a previously registered workbook-local custom function.
     * @param {string} name
     */
    unregisterFunction(name) {
        const ptr0 = passStringToWasm0(name, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.workbook_unregisterFunction(this.__wbg_ptr, ptr0, len0);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
}
if (Symbol.dispose) Workbook.prototype[Symbol.dispose] = Workbook.prototype.free;

export function init() {
    wasm.init();
}

/**
 * @param {string} formula
 * @param {FormulaDialect | null} [dialect]
 * @returns {ASTNode}
 */
export function parse(formula, dialect) {
    const ptr0 = passStringToWasm0(formula, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
    const len0 = WASM_VECTOR_LEN;
    const ret = wasm.parse(ptr0, len0, isLikeNone(dialect) ? 2 : dialect);
    if (ret[2]) {
        throw takeFromExternrefTable0(ret[1]);
    }
    return ASTNode.__wrap(ret[0]);
}

/**
 * @param {string} formula
 * @param {FormulaDialect | null} [dialect]
 * @returns {ASTNode}
 */
export function parse_formula(formula, dialect) {
    const ptr0 = passStringToWasm0(formula, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
    const len0 = WASM_VECTOR_LEN;
    const ret = wasm.parse_formula(ptr0, len0, isLikeNone(dialect) ? 2 : dialect);
    if (ret[2]) {
        throw takeFromExternrefTable0(ret[1]);
    }
    return ASTNode.__wrap(ret[0]);
}

/**
 * Recalculate formula caches while preserving all unrelated XLSX package members.
 *
 * The returned `bytes` property is a real `Uint8Array`; no base64 or integer-array
 * encoding is used. `error_location_limit` optionally caps locations per error token.
 * `options` optionally fixes the clock and RAND seed
 * (`{ rngSeed, deterministicTimestampUtc, deterministicTimezone }`); the
 * result's `clock` and `seed` replay the run. `externalLinks: 'refuse'`
 * refuses a workbook whose formulas read external link values instead of
 * using the values cached in it (`'cached'`, the default).
 * @param {Uint8Array} bytes
 * @param {number | null | undefined} error_location_limit
 * @param {any} options
 * @returns {any}
 */
export function recalculateXlsxBytes(bytes, error_location_limit, options) {
    const ret = wasm.recalculateXlsxBytes(bytes, isLikeNone(error_location_limit) ? 0x100000001 : (error_location_limit) >>> 0, options);
    if (ret[2]) {
        throw takeFromExternrefTable0(ret[1]);
    }
    return takeFromExternrefTable0(ret[0]);
}

/**
 * @param {string} formula
 * @param {FormulaDialect | null} [dialect]
 * @returns {Tokenizer}
 */
export function tokenize(formula, dialect) {
    const ptr0 = passStringToWasm0(formula, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
    const len0 = WASM_VECTOR_LEN;
    const ret = wasm.tokenize(ptr0, len0, isLikeNone(dialect) ? 2 : dialect);
    if (ret[2]) {
        throw takeFromExternrefTable0(ret[1]);
    }
    return Tokenizer.__wrap(ret[0]);
}
export function __wbg_Error_8c4e43fe74559d73(arg0, arg1) {
    const ret = Error(getStringFromWasm0(arg0, arg1));
    return ret;
}
export function __wbg_Number_04624de7d0e8332d(arg0) {
    const ret = Number(arg0);
    return ret;
}
export function __wbg_String_8f0eb39a4a4c2f66(arg0, arg1) {
    const ret = String(arg1);
    const ptr1 = passStringToWasm0(ret, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
    const len1 = WASM_VECTOR_LEN;
    getDataViewMemory0().setInt32(arg0 + 4 * 1, len1, true);
    getDataViewMemory0().setInt32(arg0 + 4 * 0, ptr1, true);
}
export function __wbg___wbindgen_bigint_get_as_i64_8fcf4ce7f1ca72a2(arg0, arg1) {
    const v = arg1;
    const ret = typeof(v) === 'bigint' ? v : undefined;
    getDataViewMemory0().setBigInt64(arg0 + 8 * 1, isLikeNone(ret) ? BigInt(0) : ret, true);
    getDataViewMemory0().setInt32(arg0 + 4 * 0, !isLikeNone(ret), true);
}
export function __wbg___wbindgen_boolean_get_bbbb1c18aa2f5e25(arg0) {
    const v = arg0;
    const ret = typeof(v) === 'boolean' ? v : undefined;
    return isLikeNone(ret) ? 0xFFFFFF : ret ? 1 : 0;
}
export function __wbg___wbindgen_debug_string_0bc8482c6e3508ae(arg0, arg1) {
    const ret = debugString(arg1);
    const ptr1 = passStringToWasm0(ret, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
    const len1 = WASM_VECTOR_LEN;
    getDataViewMemory0().setInt32(arg0 + 4 * 1, len1, true);
    getDataViewMemory0().setInt32(arg0 + 4 * 0, ptr1, true);
}
export function __wbg___wbindgen_in_47fa6863be6f2f25(arg0, arg1) {
    const ret = arg0 in arg1;
    return ret;
}
export function __wbg___wbindgen_is_bigint_31b12575b56f32fc(arg0) {
    const ret = typeof(arg0) === 'bigint';
    return ret;
}
export function __wbg___wbindgen_is_function_0095a73b8b156f76(arg0) {
    const ret = typeof(arg0) === 'function';
    return ret;
}
export function __wbg___wbindgen_is_null_ac34f5003991759a(arg0) {
    const ret = arg0 === null;
    return ret;
}
export function __wbg___wbindgen_is_object_5ae8e5880f2c1fbd(arg0) {
    const val = arg0;
    const ret = typeof(val) === 'object' && val !== null;
    return ret;
}
export function __wbg___wbindgen_is_string_cd444516edc5b180(arg0) {
    const ret = typeof(arg0) === 'string';
    return ret;
}
export function __wbg___wbindgen_is_undefined_9e4d92534c42d778(arg0) {
    const ret = arg0 === undefined;
    return ret;
}
export function __wbg___wbindgen_jsval_eq_11888390b0186270(arg0, arg1) {
    const ret = arg0 === arg1;
    return ret;
}
export function __wbg___wbindgen_jsval_loose_eq_9dd77d8cd6671811(arg0, arg1) {
    const ret = arg0 == arg1;
    return ret;
}
export function __wbg___wbindgen_number_get_8ff4255516ccad3e(arg0, arg1) {
    const obj = arg1;
    const ret = typeof(obj) === 'number' ? obj : undefined;
    getDataViewMemory0().setFloat64(arg0 + 8 * 1, isLikeNone(ret) ? 0 : ret, true);
    getDataViewMemory0().setInt32(arg0 + 4 * 0, !isLikeNone(ret), true);
}
export function __wbg___wbindgen_string_get_72fb696202c56729(arg0, arg1) {
    const obj = arg1;
    const ret = typeof(obj) === 'string' ? obj : undefined;
    var ptr1 = isLikeNone(ret) ? 0 : passStringToWasm0(ret, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
    var len1 = WASM_VECTOR_LEN;
    getDataViewMemory0().setInt32(arg0 + 4 * 1, len1, true);
    getDataViewMemory0().setInt32(arg0 + 4 * 0, ptr1, true);
}
export function __wbg___wbindgen_throw_be289d5034ed271b(arg0, arg1) {
    throw new Error(getStringFromWasm0(arg0, arg1));
}
export function __wbg_apply_ada2ee1a60ac7b3c() { return handleError(function (arg0, arg1, arg2) {
    const ret = arg0.apply(arg1, arg2);
    return ret;
}, arguments); }
export function __wbg_call_389efe28435a9388() { return handleError(function (arg0, arg1) {
    const ret = arg0.call(arg1);
    return ret;
}, arguments); }
export function __wbg_done_57b39ecd9addfe81(arg0) {
    const ret = arg0.done;
    return ret;
}
export function __wbg_entries_58c7934c745daac7(arg0) {
    const ret = Object.entries(arg0);
    return ret;
}
export function __wbg_error_7534b8e9a36f1ab4(arg0, arg1) {
    let deferred0_0;
    let deferred0_1;
    try {
        deferred0_0 = arg0;
        deferred0_1 = arg1;
        console.error(getStringFromWasm0(arg0, arg1));
    } finally {
        wasm.__wbindgen_free(deferred0_0, deferred0_1, 1);
    }
}
export function __wbg_fromEntries_7fb5bc874dbe50d5() { return handleError(function (arg0) {
    const ret = Object.fromEntries(arg0);
    return ret;
}, arguments); }
export function __wbg_from_bddd64e7d5ff6941(arg0) {
    const ret = Array.from(arg0);
    return ret;
}
export function __wbg_getTime_1e3cd1391c5c3995(arg0) {
    const ret = arg0.getTime();
    return ret;
}
export function __wbg_getTimezoneOffset_81776d10a4ec18a8(arg0) {
    const ret = arg0.getTimezoneOffset();
    return ret;
}
export function __wbg_get_9b94d73e6221f75c(arg0, arg1) {
    const ret = arg0[arg1 >>> 0];
    return ret;
}
export function __wbg_get_b3ed3ad4be2bc8ac() { return handleError(function (arg0, arg1) {
    const ret = Reflect.get(arg0, arg1);
    return ret;
}, arguments); }
export function __wbg_get_with_ref_key_1dc361bd10053bfe(arg0, arg1) {
    const ret = arg0[arg1];
    return ret;
}
export function __wbg_has_d4e53238966c12b6() { return handleError(function (arg0, arg1) {
    const ret = Reflect.has(arg0, arg1);
    return ret;
}, arguments); }
export function __wbg_instanceof_ArrayBuffer_c367199e2fa2aa04(arg0) {
    let result;
    try {
        result = arg0 instanceof ArrayBuffer;
    } catch (_) {
        result = false;
    }
    const ret = result;
    return ret;
}
export function __wbg_instanceof_Date_1b9f15b87f10aa4c(arg0) {
    let result;
    try {
        result = arg0 instanceof Date;
    } catch (_) {
        result = false;
    }
    const ret = result;
    return ret;
}
export function __wbg_instanceof_Error_8573fe0b0b480f46(arg0) {
    let result;
    try {
        result = arg0 instanceof Error;
    } catch (_) {
        result = false;
    }
    const ret = result;
    return ret;
}
export function __wbg_instanceof_Map_53af74335dec57f4(arg0) {
    let result;
    try {
        result = arg0 instanceof Map;
    } catch (_) {
        result = false;
    }
    const ret = result;
    return ret;
}
export function __wbg_instanceof_Object_1c6af87502b733ed(arg0) {
    let result;
    try {
        result = arg0 instanceof Object;
    } catch (_) {
        result = false;
    }
    const ret = result;
    return ret;
}
export function __wbg_instanceof_Uint8Array_9b9075935c74707c(arg0) {
    let result;
    try {
        result = arg0 instanceof Uint8Array;
    } catch (_) {
        result = false;
    }
    const ret = result;
    return ret;
}
export function __wbg_isArray_d314bb98fcf08331(arg0) {
    const ret = Array.isArray(arg0);
    return ret;
}
export function __wbg_isSafeInteger_bfbc7332a9768d2a(arg0) {
    const ret = Number.isSafeInteger(arg0);
    return ret;
}
export function __wbg_iterator_6ff6560ca1568e55() {
    const ret = Symbol.iterator;
    return ret;
}
export function __wbg_keys_b50a709a76add04e(arg0) {
    const ret = Object.keys(arg0);
    return ret;
}
export function __wbg_length_32ed9a279acd054c(arg0) {
    const ret = arg0.length;
    return ret;
}
export function __wbg_length_35a7bace40f36eac(arg0) {
    const ret = arg0.length;
    return ret;
}
export function __wbg_message_9ddc4b9a62a7c379(arg0) {
    const ret = arg0.message;
    return ret;
}
export function __wbg_name_446e25ef2cfdab5a(arg0) {
    const ret = arg0.name;
    return ret;
}
export function __wbg_new_0_73afc35eb544e539() {
    const ret = new Date();
    return ret;
}
export function __wbg_new_245cd5c49157e602(arg0) {
    const ret = new Date(arg0);
    return ret;
}
export function __wbg_new_361308b2356cecd0() {
    const ret = new Object();
    return ret;
}
export function __wbg_new_3eb36ae241fe6f44() {
    const ret = new Array();
    return ret;
}
export function __wbg_new_72b49615380db768(arg0, arg1) {
    const ret = new Error(getStringFromWasm0(arg0, arg1));
    return ret;
}
export function __wbg_new_8a6f238a6ece86ea() {
    const ret = new Error();
    return ret;
}
export function __wbg_new_dca287b076112a51() {
    const ret = new Map();
    return ret;
}
export function __wbg_new_dd2b680c8bf6ae29(arg0) {
    const ret = new Uint8Array(arg0);
    return ret;
}
export function __wbg_new_from_slice_a3d2629dc1826784(arg0, arg1) {
    const ret = new Uint8Array(getArrayU8FromWasm0(arg0, arg1));
    return ret;
}
export function __wbg_new_no_args_1c7c842f08d00ebb(arg0, arg1) {
    const ret = new Function(getStringFromWasm0(arg0, arg1));
    return ret;
}
export function __wbg_next_3482f54c49e8af19() { return handleError(function (arg0) {
    const ret = arg0.next();
    return ret;
}, arguments); }
export function __wbg_next_418f80d8f5303233(arg0) {
    const ret = arg0.next;
    return ret;
}
export function __wbg_now_2c95c9de01293173(arg0) {
    const ret = arg0.now();
    return ret;
}
export function __wbg_ownKeys_c7100fb5fa376c6f() { return handleError(function (arg0) {
    const ret = Reflect.ownKeys(arg0);
    return ret;
}, arguments); }
export function __wbg_performance_7a3ffd0b17f663ad(arg0) {
    const ret = arg0.performance;
    return ret;
}
export function __wbg_prototypesetcall_bdcdcc5842e4d77d(arg0, arg1, arg2) {
    Uint8Array.prototype.set.call(getArrayU8FromWasm0(arg0, arg1), arg2);
}
export function __wbg_push_8ffdcb2063340ba5(arg0, arg1) {
    const ret = arg0.push(arg1);
    return ret;
}
export function __wbg_set_1eb0999cf5d27fc8(arg0, arg1, arg2) {
    const ret = arg0.set(arg1, arg2);
    return ret;
}
export function __wbg_set_3f1d0b984ed272ed(arg0, arg1, arg2) {
    arg0[arg1] = arg2;
}
export function __wbg_set_6cb8631f80447a67() { return handleError(function (arg0, arg1, arg2) {
    const ret = Reflect.set(arg0, arg1, arg2);
    return ret;
}, arguments); }
export function __wbg_set_f43e577aea94465b(arg0, arg1, arg2) {
    arg0[arg1 >>> 0] = arg2;
}
export function __wbg_stack_0ed75d68575b0f3c(arg0, arg1) {
    const ret = arg1.stack;
    const ptr1 = passStringToWasm0(ret, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
    const len1 = WASM_VECTOR_LEN;
    getDataViewMemory0().setInt32(arg0 + 4 * 1, len1, true);
    getDataViewMemory0().setInt32(arg0 + 4 * 0, ptr1, true);
}
export function __wbg_static_accessor_GLOBAL_12837167ad935116() {
    const ret = typeof global === 'undefined' ? null : global;
    return isLikeNone(ret) ? 0 : addToExternrefTable0(ret);
}
export function __wbg_static_accessor_GLOBAL_THIS_e628e89ab3b1c95f() {
    const ret = typeof globalThis === 'undefined' ? null : globalThis;
    return isLikeNone(ret) ? 0 : addToExternrefTable0(ret);
}
export function __wbg_static_accessor_SELF_a621d3dfbb60d0ce() {
    const ret = typeof self === 'undefined' ? null : self;
    return isLikeNone(ret) ? 0 : addToExternrefTable0(ret);
}
export function __wbg_static_accessor_WINDOW_f8727f0cf888e0bd() {
    const ret = typeof window === 'undefined' ? null : window;
    return isLikeNone(ret) ? 0 : addToExternrefTable0(ret);
}
export function __wbg_stringify_8d1cc6ff383e8bae() { return handleError(function (arg0) {
    const ret = JSON.stringify(arg0);
    return ret;
}, arguments); }
export function __wbg_toString_3cadee6e7c22b39e() { return handleError(function (arg0, arg1) {
    const ret = arg0.toString(arg1);
    return ret;
}, arguments); }
export function __wbg_value_0546255b415e96c1(arg0) {
    const ret = arg0.value;
    return ret;
}
export function __wbindgen_cast_0000000000000001(arg0) {
    // Cast intrinsic for `F64 -> Externref`.
    const ret = arg0;
    return ret;
}
export function __wbindgen_cast_0000000000000002(arg0) {
    // Cast intrinsic for `I64 -> Externref`.
    const ret = arg0;
    return ret;
}
export function __wbindgen_cast_0000000000000003(arg0, arg1) {
    // Cast intrinsic for `Ref(String) -> Externref`.
    const ret = getStringFromWasm0(arg0, arg1);
    return ret;
}
export function __wbindgen_cast_0000000000000004(arg0) {
    // Cast intrinsic for `U64 -> Externref`.
    const ret = BigInt.asUintN(64, arg0);
    return ret;
}
export function __wbindgen_init_externref_table() {
    const table = wasm.__wbindgen_externrefs;
    const offset = table.grow(4);
    table.set(0, undefined);
    table.set(offset + 0, undefined);
    table.set(offset + 1, null);
    table.set(offset + 2, true);
    table.set(offset + 3, false);
}
export function __wbindgen_object_is_undefined(arg0) {
    const ret = arg0 === undefined;
    return ret;
}
const ASTNodeFinalization = (typeof FinalizationRegistry === 'undefined')
    ? { register: () => {}, unregister: () => {} }
    : new FinalizationRegistry(ptr => wasm.__wbg_astnode_free(ptr >>> 0, 1));
const JsTokenFinalization = (typeof FinalizationRegistry === 'undefined')
    ? { register: () => {}, unregister: () => {} }
    : new FinalizationRegistry(ptr => wasm.__wbg_jstoken_free(ptr >>> 0, 1));
const ParserFinalization = (typeof FinalizationRegistry === 'undefined')
    ? { register: () => {}, unregister: () => {} }
    : new FinalizationRegistry(ptr => wasm.__wbg_parser_free(ptr >>> 0, 1));
const ParserErrorFinalization = (typeof FinalizationRegistry === 'undefined')
    ? { register: () => {}, unregister: () => {} }
    : new FinalizationRegistry(ptr => wasm.__wbg_parsererror_free(ptr >>> 0, 1));
const ReferenceFinalization = (typeof FinalizationRegistry === 'undefined')
    ? { register: () => {}, unregister: () => {} }
    : new FinalizationRegistry(ptr => wasm.__wbg_reference_free(ptr >>> 0, 1));
const SheetFinalization = (typeof FinalizationRegistry === 'undefined')
    ? { register: () => {}, unregister: () => {} }
    : new FinalizationRegistry(ptr => wasm.__wbg_sheet_free(ptr >>> 0, 1));
const SheetPortSessionFinalization = (typeof FinalizationRegistry === 'undefined')
    ? { register: () => {}, unregister: () => {} }
    : new FinalizationRegistry(ptr => wasm.__wbg_sheetportsession_free(ptr >>> 0, 1));
const TokenizerFinalization = (typeof FinalizationRegistry === 'undefined')
    ? { register: () => {}, unregister: () => {} }
    : new FinalizationRegistry(ptr => wasm.__wbg_tokenizer_free(ptr >>> 0, 1));
const TokenizerErrorFinalization = (typeof FinalizationRegistry === 'undefined')
    ? { register: () => {}, unregister: () => {} }
    : new FinalizationRegistry(ptr => wasm.__wbg_tokenizererror_free(ptr >>> 0, 1));
const WorkbookFinalization = (typeof FinalizationRegistry === 'undefined')
    ? { register: () => {}, unregister: () => {} }
    : new FinalizationRegistry(ptr => wasm.__wbg_workbook_free(ptr >>> 0, 1));

function addToExternrefTable0(obj) {
    const idx = wasm.__externref_table_alloc();
    wasm.__wbindgen_externrefs.set(idx, obj);
    return idx;
}

function _assertClass(instance, klass) {
    if (!(instance instanceof klass)) {
        throw new Error(`expected instance of ${klass.name}`);
    }
}

function debugString(val) {
    // primitive types
    const type = typeof val;
    if (type == 'number' || type == 'boolean' || val == null) {
        return  `${val}`;
    }
    if (type == 'string') {
        return `"${val}"`;
    }
    if (type == 'symbol') {
        const description = val.description;
        if (description == null) {
            return 'Symbol';
        } else {
            return `Symbol(${description})`;
        }
    }
    if (type == 'function') {
        const name = val.name;
        if (typeof name == 'string' && name.length > 0) {
            return `Function(${name})`;
        } else {
            return 'Function';
        }
    }
    // objects
    if (Array.isArray(val)) {
        const length = val.length;
        let debug = '[';
        if (length > 0) {
            debug += debugString(val[0]);
        }
        for(let i = 1; i < length; i++) {
            debug += ', ' + debugString(val[i]);
        }
        debug += ']';
        return debug;
    }
    // Test for built-in
    const builtInMatches = /\[object ([^\]]+)\]/.exec(toString.call(val));
    let className;
    if (builtInMatches && builtInMatches.length > 1) {
        className = builtInMatches[1];
    } else {
        // Failed to match the standard '[object ClassName]'
        return toString.call(val);
    }
    if (className == 'Object') {
        // we're a user defined class or Object
        // JSON.stringify avoids problems with cycles, and is generally much
        // easier than looping through ownProperties of `val`.
        try {
            return 'Object(' + JSON.stringify(val) + ')';
        } catch (_) {
            return 'Object';
        }
    }
    // errors
    if (val instanceof Error) {
        return `${val.name}: ${val.message}\n${val.stack}`;
    }
    // TODO we could test for more things here, like `Set`s and `Map`s.
    return className;
}

function getArrayU8FromWasm0(ptr, len) {
    ptr = ptr >>> 0;
    return getUint8ArrayMemory0().subarray(ptr / 1, ptr / 1 + len);
}

let cachedDataViewMemory0 = null;
function getDataViewMemory0() {
    if (cachedDataViewMemory0 === null || cachedDataViewMemory0.buffer.detached === true || (cachedDataViewMemory0.buffer.detached === undefined && cachedDataViewMemory0.buffer !== wasm.memory.buffer)) {
        cachedDataViewMemory0 = new DataView(wasm.memory.buffer);
    }
    return cachedDataViewMemory0;
}

function getStringFromWasm0(ptr, len) {
    ptr = ptr >>> 0;
    return decodeText(ptr, len);
}

let cachedUint8ArrayMemory0 = null;
function getUint8ArrayMemory0() {
    if (cachedUint8ArrayMemory0 === null || cachedUint8ArrayMemory0.byteLength === 0) {
        cachedUint8ArrayMemory0 = new Uint8Array(wasm.memory.buffer);
    }
    return cachedUint8ArrayMemory0;
}

function handleError(f, args) {
    try {
        return f.apply(this, args);
    } catch (e) {
        const idx = addToExternrefTable0(e);
        wasm.__wbindgen_exn_store(idx);
    }
}

function isLikeNone(x) {
    return x === undefined || x === null;
}

function passArray8ToWasm0(arg, malloc) {
    const ptr = malloc(arg.length * 1, 1) >>> 0;
    getUint8ArrayMemory0().set(arg, ptr / 1);
    WASM_VECTOR_LEN = arg.length;
    return ptr;
}

function passStringToWasm0(arg, malloc, realloc) {
    if (realloc === undefined) {
        const buf = cachedTextEncoder.encode(arg);
        const ptr = malloc(buf.length, 1) >>> 0;
        getUint8ArrayMemory0().subarray(ptr, ptr + buf.length).set(buf);
        WASM_VECTOR_LEN = buf.length;
        return ptr;
    }

    let len = arg.length;
    let ptr = malloc(len, 1) >>> 0;

    const mem = getUint8ArrayMemory0();

    let offset = 0;

    for (; offset < len; offset++) {
        const code = arg.charCodeAt(offset);
        if (code > 0x7F) break;
        mem[ptr + offset] = code;
    }
    if (offset !== len) {
        if (offset !== 0) {
            arg = arg.slice(offset);
        }
        ptr = realloc(ptr, len, len = offset + arg.length * 3, 1) >>> 0;
        const view = getUint8ArrayMemory0().subarray(ptr + offset, ptr + len);
        const ret = cachedTextEncoder.encodeInto(arg, view);

        offset += ret.written;
        ptr = realloc(ptr, len, offset, 1) >>> 0;
    }

    WASM_VECTOR_LEN = offset;
    return ptr;
}

function takeFromExternrefTable0(idx) {
    const value = wasm.__wbindgen_externrefs.get(idx);
    wasm.__externref_table_dealloc(idx);
    return value;
}

let cachedTextDecoder = new TextDecoder('utf-8', { ignoreBOM: true, fatal: true });
cachedTextDecoder.decode();
const MAX_SAFARI_DECODE_BYTES = 2146435072;
let numBytesDecoded = 0;
function decodeText(ptr, len) {
    numBytesDecoded += len;
    if (numBytesDecoded >= MAX_SAFARI_DECODE_BYTES) {
        cachedTextDecoder = new TextDecoder('utf-8', { ignoreBOM: true, fatal: true });
        cachedTextDecoder.decode();
        numBytesDecoded = len;
    }
    return cachedTextDecoder.decode(getUint8ArrayMemory0().subarray(ptr, ptr + len));
}

const cachedTextEncoder = new TextEncoder();

if (!('encodeInto' in cachedTextEncoder)) {
    cachedTextEncoder.encodeInto = function (arg, view) {
        const buf = cachedTextEncoder.encode(arg);
        view.set(buf);
        return {
            read: arg.length,
            written: buf.length
        };
    };
}

let WASM_VECTOR_LEN = 0;


let wasm;
export function __wbg_set_wasm(val) {
    wasm = val;
}
