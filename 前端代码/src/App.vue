<script setup>
import { ref, reactive, computed } from 'vue'
import axios from 'axios'
import { ElMessage } from 'element-plus'
import { UploadFilled, Setting, Download, QuestionFilled } from '@element-plus/icons-vue'

// 后端地址（如需跨机部署，改为你的服务器地址）
const API = axios.create({ baseURL: 'http://127.0.0.1:5000' })

const uploading = ref(false)
const fileId = ref('')
const sheets = ref([])
const selectedSheet = ref('')
const sniff = ref({})
const columns = ref([])
const timeCandidates = ref([])
const numericCols = ref([])

const plotType = ref('timeseries') // 'timeseries' | 'hysteresis'
const createDefaultParams = () => ({
    // 通用
    title: '',
    xlabel: '',
    ylabel: '',
    legend_loc: 'upper right',
    figsize_cm: [16, 9],
    dpi: 120,
    x_major: null,
    y_major: null,
    show_minor_grid: true,
    linewidth: 2.0,

    // 时程
    time_col: '',
    series_cols: [],
    labels: [],

    // 滞回
    x_col: '',
    y_cols: [],

    // 附加
    zero_baseline: true,
    zero_axes: true,
    equal_aspect: false,
    annotate_peaks: true,
    title_pad: 10,

    metrics_box: false,

    // 坐标轴范围控制
    xlim: [null, null],
    ylim: [null, null],

    // 平滑
    smooth: 'savgol', // 'savgol' | 'ma' | 'none'
    smooth_kwargs: { window_length: 11, polyorder: 3, k: 5 },

    // 导出
    export_formats: ['png', 'pdf', 'svg'],
    filename_base: '',
    save_dir: ''
})
const params = reactive(createDefaultParams())

const previewUrl = ref('')
const previewLoading = ref(false)
const activeTab = ref('upload')
const showHelp = ref(false)

function resetParams() {
    Object.assign(params, createDefaultParams())
    ElMessage.success('已重置参数')
}

function clearPreview() {
    previewUrl.value = ''
    downloadLinks.value = []
    zipLink.value = ''
    ElMessage.success('已清空预览与下载链接')
}

async function onUpload(file) {
    try {
        uploading.value = true
        const fd = new FormData()
        fd.append('file', file.raw)
        const { data } = await API.post('/api/excel/upload', fd, {
            headers: { 'Content-Type': 'multipart/form-data' }
        })
        if (!data.ok) return ElMessage.error(data.msg || '上传失败')
        fileId.value = data.file_id
        sheets.value = data.sheets
        sniff.value = data.sniff || {}
        selectedSheet.value = sheets.value[0] || ''
        ElMessage.success('上传成功')
        await fetchColumns()
    } catch (e) {
        ElMessage.error('上传失败：' + e)
    } finally {
        uploading.value = false
    }
}

async function fetchColumns() {
    if (!fileId.value || !selectedSheet.value) return
    const { data } = await API.get('/api/excel/columns', {
        params: { file_id: fileId.value, sheet: selectedSheet.value }
    })
    if (!data.ok) return ElMessage.error(data.msg || '列读取失败')
    columns.value = data.columns || []
    timeCandidates.value = data.time_candidates || []
    numericCols.value = data.numeric_cols || []
    // 智能填充
    if (!params.time_col && timeCandidates.value.length) {
        params.time_col = timeCandidates.value[0]
    }
    if (plotType.value === 'timeseries' && params.series_cols.length === 0) {
        params.series_cols = numericCols.value.filter(c => c !== params.time_col).slice(0, 2)
        params.labels = [...params.series_cols]
    }
    if (plotType.value === 'hysteresis' && !params.x_col) {
        params.x_col = numericCols.value[0] || ''
        params.y_cols = numericCols.value.slice(1, 3)
    }
}

function labelPlaceholder() {
    return plotType.value === 'timeseries' ? '与 series_cols 对应顺序' : '与 y_cols 对应顺序'
}

function processParams(p) {
    // 处理坐标轴范围：将 null 值保持为 null，有效数字转换为数组
    const processLim = (lim) => {
        if (!Array.isArray(lim) || lim.length !== 2) return null
        const processed = lim.map(v => v === null || v === undefined || v === '' ? null : v)
        return processed.some(v => v !== null) ? processed : null
    }

    if (p.xlim) p.xlim = processLim(p.xlim)
    if (p.ylim) p.ylim = processLim(p.ylim)
    return p
}

async function doPreview() {
    if (!fileId.value) return ElMessage.warning('请先上传 Excel')
    if (!selectedSheet.value) return ElMessage.warning('请选择 Sheet')

    const p = processParams(JSON.parse(JSON.stringify(params)))
    p.dpi = 120 // 预览低 DPI
    try {
        previewLoading.value = true
        const { data } = await API.post('/api/plot/preview', {
            file_id: fileId.value,
            sheet: selectedSheet.value,
            plot_type: plotType.value,
            params: p
        })
        if (!data.ok) return ElMessage.error(data.msg || '预览失败')
        previewUrl.value = data.preview_data_url
    } finally {
        previewLoading.value = false
    }
}

async function doExport() {
    if (!fileId.value) return ElMessage.warning('请先上传 Excel')
    const p = processParams(JSON.parse(JSON.stringify(params)))
    p.dpi = 600
    if (!p.filename_base) {
        const ts = new Date().toISOString().slice(0, 19).replace(/[:T]/g, '')
        p.filename_base = (plotType.value === 'timeseries' ? 'timeseries_' : 'hysteresis_') + ts
    }
    const { data } = await API.post('/api/plot/export', {
        file_id: fileId.value,
        sheet: selectedSheet.value,
        plot_type: plotType.value,
        params: p
    })
    if (!data.ok) return ElMessage.error(data.msg || '导出失败')
    // 显示下载链接
    downloadLinks.value = data.files || []
    zipLink.value = data.zip || ''
    previewUrl.value = data.preview_data_url || previewUrl.value
    ElMessage.success('导出完成')
}

const downloadLinks = ref([])
const zipLink = ref('')
</script>

<template>
    <div class="app-container">
        <div class="app-hero">
            <div class="title-wrap">
                <h1 class="app-title">科研绘图助手 · MVP</h1>
                <p class="app-subtitle">读取 Excel → 参数试配（预览）→ 导出 PNG/PDF/SVG</p>
            </div>
            <div class="header-actions">
                <el-button type="primary" plain size="small" @click="showHelp = true">
                    <el-icon style="margin-right:4px">
                        <QuestionFilled />
                    </el-icon>查看帮助
                </el-button>
                <el-button size="small" @click="resetParams">重置参数</el-button>
                <el-button size="small" @click="clearPreview">清空预览</el-button>
            </div>
        </div>

        <el-tabs v-model="activeTab" class="main-tabs" stretch>
            <el-tab-pane name="upload">
                <template #label>
                    <span class="tab-label"><el-icon>
                            <UploadFilled />
                        </el-icon><span>上传 Excel</span></span>
                </template>
                <el-card shadow="never">
                    <el-upload drag :auto-upload="false" accept=".xlsx,.xls" :on-change="onUpload" :disabled="uploading"
                        style="width:100%;">
                        <el-icon class="el-icon--upload">
                            <UploadFilled />
                        </el-icon>
                        <div class="el-upload__text">将文件拖到此处，或 <em>点击上传</em></div>
                        <template #tip>
                            <div class="el-upload__tip">仅支持 .xlsx/.xls</div>
                        </template>
                    </el-upload>
                    <div v-if="fileId" style="margin-top:12px;">
                        <el-form :inline="true">
                            <el-form-item label="选择 Sheet">
                                <el-select v-model="selectedSheet" placeholder="选择 Sheet" @change="fetchColumns"
                                    style="width:240px;">
                                    <el-option v-for="s in sheets" :key="s" :label="s" :value="s" />
                                </el-select>
                            </el-form-item>
                        </el-form>
                    </div>
                </el-card>
            </el-tab-pane>

            <el-tab-pane name="config">
                <template #label>
                    <span class="tab-label"><el-icon>
                            <Setting />
                        </el-icon><span>参数试配</span></span>
                </template>
                <el-card shadow="never">
                    <el-radio-group v-model="plotType" size="small" @change="fetchColumns" style="margin-bottom:12px;">
                        <el-radio-button label="timeseries">时程曲线</el-radio-button>
                        <el-radio-button label="hysteresis">滞回曲线</el-radio-button>
                    </el-radio-group>

                    <el-row :gutter="16">
                        <el-col :span="12">
                            <el-form label-width="120px" size="small">
                                <template v-if="plotType === 'timeseries'">
                                    <el-form-item label="时间列 time_col">
                                        <el-select v-model="params.time_col" filterable style="width:100%;">
                                            <el-option v-for="c in columns" :key="c" :label="c" :value="c" />
                                        </el-select>
                                    </el-form-item>

                                    <el-form-item label="series_cols（多选）">
                                        <el-select v-model="params.series_cols" multiple filterable style="width:100%;">
                                            <el-option v-for="c in numericCols" :key="c" :label="c" :value="c" />
                                        </el-select>
                                    </el-form-item>

                                    <el-form-item label="labels（可选）" :title="labelPlaceholder()">
                                        <el-input v-model="params.labels" type="textarea" :rows="2"
                                            placeholder="按逗号分隔，与 series_cols 顺序一一对应"
                                            @change="() => { if (typeof params.labels === 'string') params.labels = params.labels.split(',').map(s => s.trim()) }" />
                                    </el-form-item>

                                    <el-form-item label="零基线">
                                        <el-switch v-model="params.zero_baseline" />
                                    </el-form-item>

                                    <el-form-item label="峰值标注">
                                        <el-switch v-model="params.annotate_peaks" />
                                    </el-form-item>

                                    <el-form-item label="峰值对比框">
                                        <el-switch v-model="params.metrics_box" />
                                        <span style="margin-left:8px;color:#999;font-size:12px;">仅两条曲线时生效</span>
                                    </el-form-item>
                                </template>

                                <template v-else>
                                    <el-form-item label="位移列 x_col">
                                        <el-select v-model="params.x_col" filterable style="width:100%;">
                                            <el-option v-for="c in numericCols" :key="c" :label="c" :value="c" />
                                        </el-select>
                                    </el-form-item>

                                    <el-form-item label="力列 y_cols（多选）">
                                        <el-select v-model="params.y_cols" multiple filterable style="width:100%;">
                                            <el-option v-for="c in numericCols" :key="c" :label="c" :value="c" />
                                        </el-select>
                                    </el-form-item>

                                    <el-form-item label="labels（可选）" :title="labelPlaceholder()">
                                        <el-input v-model="params.labels" type="textarea" :rows="2"
                                            placeholder="按逗号分隔，与 y_cols 顺序一一对应"
                                            @change="() => { if (typeof params.labels === 'string') params.labels = params.labels.split(',').map(s => s.trim()) }" />
                                    </el-form-item>

                                    <el-form-item label="零轴 & 等比例">
                                        <el-switch v-model="params.zero_axes" style="margin-right:12px;" />零轴
                                        <el-switch v-model="params.equal_aspect" style="margin-left:24px;" />等比例
                                    </el-form-item>

                                    <el-form-item label="峰值标注">
                                        <el-switch v-model="params.annotate_peaks" />
                                    </el-form-item>
                                </template>

                                <el-form-item label="标题 / 轴标签">
                                    <el-input v-model="params.title" placeholder="标题"
                                        style="width:33%;margin-right:8px;" />
                                    <el-input v-model="params.xlabel" placeholder="X 轴"
                                        style="width:30%;margin-right:8px;" />
                                    <el-input v-model="params.ylabel" placeholder="Y 轴" style="width:30%;" />
                                </el-form-item>

                                <el-form-item label="标题边距">
                                    <el-input-number v-model="params.title_pad" :min="0" :max="60" />
                                </el-form-item>

                                <el-form-item label="图幅(cm)">
                                    <el-input-number v-model="params.figsize_cm[0]" :min="6" :max="60" /> ×
                                    <el-input-number v-model="params.figsize_cm[1]" :min="6" :max="60" />
                                </el-form-item>

                                <el-form-item label="主刻度">
                                    <el-input-number v-model="params.x_major" :min="0" :step="0.5" placeholder="X 步长" />
                                    /
                                    <el-input-number v-model="params.y_major" :min="0" :step="0.5" placeholder="Y 步长" />
                                    <el-checkbox v-model="params.show_minor_grid"
                                        style="margin-left:12px;">次网格</el-checkbox>
                                </el-form-item>

                                <el-form-item label="X轴范围">
                                    <el-input-number v-model="params.xlim[0]" placeholder="最小值" style="width:120px;" />
                                    <span style="margin:0 8px;">至</span>
                                    <el-input-number v-model="params.xlim[1]" placeholder="最大值" style="width:120px;" />
                                    <span style="margin-left:8px;color:#999;font-size:12px;">留空自适应</span>
                                </el-form-item>

                                <el-form-item label="Y轴范围">
                                    <el-input-number v-model="params.ylim[0]" placeholder="最小值" style="width:120px;" />
                                    <span style="margin:0 8px;">至</span>
                                    <el-input-number v-model="params.ylim[1]" placeholder="最大值" style="width:120px;" />
                                    <span style="margin-left:8px;color:#999;font-size:12px;">留空自适应</span>
                                </el-form-item>

                                <el-form-item label="图例位置">
                                    <el-select v-model="params.legend_loc" style="width:240px;">
                                        <el-option
                                            v-for="loc in ['upper right', 'upper left', 'lower right', 'lower left', 'best', 'center', 'center left', 'center right']"
                                            :key="loc" :label="loc" :value="loc" />
                                    </el-select>
                                </el-form-item>

                                <el-form-item label="线宽 / 平滑">
                                    <el-input-number v-model="params.linewidth" :min="0.5" :max="6" :step="0.1" />
                                    <el-select v-model="params.smooth" style="width:160px; margin-left:12px;">
                                        <el-option label="Savitzky-Golay" value="savgol" />
                                        <el-option label="滑动平均" value="ma" />
                                        <el-option label="无" value="none" />
                                    </el-select>
                                </el-form-item>


                            </el-form>
                        </el-col>

                        <el-col :span="12">
                            <el-card class="preview-card" shadow="never" :body-style="{ padding: '12px' }"
                                v-loading="previewLoading">
                                <div class="preview-toolbar">
                                    <div>
                                        <el-button type="primary" size="small" @click="doPreview">刷新预览</el-button>
                                        <el-tag type="info" size="small" class="ml8">预览 DPI: 120</el-tag>
                                    </div>
                                </div>
                                <div class="preview-box">
                                    <img v-if="previewUrl" :src="previewUrl" alt="预览图" class="preview-img" />
                                    <el-empty v-else description="暂无预览，请点击“刷新预览”" />
                                </div>
                            </el-card>
                        </el-col>
                    </el-row>
                </el-card>
            </el-tab-pane>

            <el-tab-pane name="export">
                <template #label>
                    <span class="tab-label"><el-icon>
                            <Download />
                        </el-icon><span>导出与下载</span></span>
                </template>
                <el-card shadow="never">
                    <template #header>③ 导出 & 下载</template>
                    <el-form label-width="120px" size="small">
                        <el-form-item label="导出格式">
                            <el-checkbox-group v-model="params.export_formats">
                                <el-checkbox label="png" />
                                <el-checkbox label="pdf" />
                                <el-checkbox label="svg" />
                            </el-checkbox-group>
                        </el-form-item>
                        <el-form-item label="文件名 / 目录">
                            <el-input v-model="params.filename_base" placeholder="留空将自动命名"
                                style="width:300px; margin-right:8px;" />
                            <el-input v-model="params.save_dir" placeholder="保存目录（可选）" style="width:360px;" />
                        </el-form-item>
                        <el-form-item>
                            <el-button type="success" @click="doExport">导出高清图</el-button>
                            <el-tag type="success" size="small" style="margin-left:8px;">导出 DPI: 600</el-tag>
                        </el-form-item>
                    </el-form>

                    <div v-if="downloadLinks.length || zipLink" class="download-wrap">
                        <p class="download-title">下载链接：</p>
                        <ul class="download-list">
                            <li v-for="(u, idx) in downloadLinks" :key="idx">
                                <a :href="API.defaults.baseURL + u" target="_blank">{{ u.split('=').pop() }}</a>
                            </li>
                        </ul>
                        <p v-if="zipLink" class="download-zip">打包下载：
                            <a :href="API.defaults.baseURL + zipLink" target="_blank">{{ zipLink.split('=').pop() }}</a>
                        </p>
                    </div>
                </el-card>
            </el-tab-pane>
        </el-tabs>

        <!-- 帮助对话框 -->
        <el-dialog v-model="showHelp" title="使用帮助" width="920px" top="6vh" :append-to-body="true">
            <div class="help-content">
                <div class="help-intro">
                    <p>本工具用于将 Excel 数据快速绘制为科研图，支持时程曲线与滞回曲线，提供峰值标注、网格控制、坐标轴范围控制及多格式高清导出（PNG/PDF/SVG）。</p>
                </div>
                <div class="help-grid">
                    <section class="help-section">
                        <h3>① 上传 Excel</h3>
                        <ul>
                            <li>支持 .xlsx / .xls；上传后选择 Sheet。</li>
                            <li>系统会智能识别时间列与数值列，供下步选择。</li>
                        </ul>
                    </section>
                    <section class="help-section">
                        <h3>② 参数试配</h3>
                        <ul>
                            <li><b>时程曲线</b>：选择 <code>time_col</code> 与多条 <code>series_cols</code>，可选自定义
                                <code>labels</code>。
                            </li>
                            <li><b>滞回曲线</b>：选择位移列 <code>x_col</code> 与一到多条力列 <code>y_cols</code>。</li>
                            <li>可设置主刻度步长、是否显示次网格、平滑方式与线宽。</li>
                            <li>坐标范围 <code>xlim/ylim</code> 支持部分留空（自适应）。</li>
                            <li>启用“峰值标注/峰值对比框”可自动标注并避免重叠。</li>
                        </ul>
                    </section>
                    <section class="help-section">
                        <h3>③ 预览与导出</h3>
                        <ul>
                            <li>右侧点击“刷新预览”生成预览图（低 DPI）。</li>
                            <li>在“导出与下载”中选择格式并导出（高 DPI）。</li>
                            <li>导出完成后会出现单文件链接与打包下载链接。</li>
                        </ul>
                    </section>
                    <section class="help-section">
                        <h3>风格建议</h3>
                        <ul>
                            <li>建议标题简洁、轴标签含单位；图例放置在不遮挡区域。</li>
                            <li>若需要标准论文风格，可在后端调整 <code>utils_plot.py</code> 的坐标轴与网格参数。</li>
                            <li>刻度、网格、线宽请依据期刊模板微调。</li>
                        </ul>
                    </section>
                    <section class="help-section">
                        <h3>常见问题</h3>
                        <ul>
                            <li>预览为空：请确认已选择 Sheet 与必要列。</li>
                            <li>文字显示异常：确保本机装有中文字体（如“微软雅黑/思源黑体”）。</li>
                            <li>导出失败：检查保存目录权限或更换导出格式再试。</li>
                        </ul>
                    </section>
                </div>
            </div>
            <template #footer>
                <span class="dialog-footer">
                    <el-button type="primary" @click="showHelp = false">已了解</el-button>
                </span>
            </template>
        </el-dialog>
    </div>
</template>

<style>
body {
    background: #fafafa;
}

.app-container {
    max-width: 1200px;
    margin: 0 auto;
    padding: 20px 24px;
}

.app-hero {
    display: flex;
    align-items: flex-end;
    justify-content: space-between;
    margin-bottom: 12px;
    padding: 14px 16px;
    border: 1px solid #eee;
    border-radius: 12px;
    background: linear-gradient(135deg, #f2f8ff 0%, #ffffff 45%, #fff7f2 100%);
    box-shadow: 0 4px 18px rgba(0, 0, 0, 0.04);
}

.app-title {
    margin: 0 0 4px 0;
    font-size: 24px;
    font-weight: 700;
}

.app-subtitle {
    margin: 0;
    color: #666;
    font-size: 14.5px;
}

.main-tabs {
    margin-top: 8px;
}

.main-tabs .el-tabs__item {
    font-size: 15px;
    font-weight: 600;
}

.tab-label {
    display: inline-flex;
    align-items: center;
    gap: 6px;
}

/* 帮助弹窗样式 */
.help-content {
    line-height: 1.8;
    font-size: 14px;
}

.help-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 12px 24px;
}

.help-section h3 {
    margin: 8px 0 6px;
    font-size: 16px;
}

.help-section ul {
    margin: 0 0 4px 16px;
    padding: 0;
}

.help-intro {
    margin-bottom: 6px;
}

@media (max-width: 960px) {
    .help-grid {
        grid-template-columns: 1fr;
    }
}

.header-actions .el-button+.el-button {
    margin-left: 8px;
}

.preview-card {
    position: sticky;
    top: 12px;
    border-radius: 10px;
}

.preview-toolbar {
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 8px;
}

.ml8 {
    margin-left: 8px;
}

.preview-box {
    min-height: 340px;
    display: flex;
    align-items: center;
    justify-content: center;
}

.preview-img {
    max-width: 100%;
    max-height: 100%;
    border-radius: 6px;
}

.download-wrap {
    margin-top: 8px;
}

.download-title {
    margin: 0 0 6px;
}

.download-list {
    margin: 0 0 8px 16px;
    padding: 0;
}

.download-list li {
    list-style: disc;
    margin-bottom: 4px;
}

.download-list a {
    color: #409eff;
    text-decoration: none;
}

.download-list a:hover {
    text-decoration: underline;
}
</style>
