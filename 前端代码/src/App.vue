<script setup>
import { ref, reactive, computed } from 'vue'
import axios from 'axios'
import { ElMessage } from 'element-plus'
import { UploadFilled } from '@element-plus/icons-vue'

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
    show_indicator: true,

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

async function doPreview() {
    if (!fileId.value) return ElMessage.warning('请先上传 Excel')
    if (!selectedSheet.value) return ElMessage.warning('请选择 Sheet')

    const p = JSON.parse(JSON.stringify(params))
    p.dpi = 120 // 预览低 DPI
    const { data } = await API.post('/api/plot/preview', {
        file_id: fileId.value,
        sheet: selectedSheet.value,
        plot_type: plotType.value,
        params: p
    })
    if (!data.ok) return ElMessage.error(data.msg || '预览失败')
    previewUrl.value = data.preview_data_url
}

async function doExport() {
    if (!fileId.value) return ElMessage.warning('请先上传 Excel')
    const p = JSON.parse(JSON.stringify(params))
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
    <div style="max-width:1200px;margin:0 auto;padding:24px;">
        <h2 style="margin-bottom:8px;">科研绘图助手 · MVP</h2>
        <p style="margin-top:0;color:#666;">读取 Excel → 参数试配（预览）→ 导出 PNG/PDF/SVG</p>
        <div style="margin:8px 0 16px;">
            <el-button size="small" @click="resetParams">重置参数</el-button>
            <el-button size="small" @click="clearPreview" style="margin-left:8px;">清空预览</el-button>
        </div>

        <!-- Step 1: 上传与 Sheet 选择 -->
        <el-card shadow="never" style="margin-bottom:16px;">
            <template #header>① 上传 Excel</template>
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

        <!-- Step 2: 参数配置 -->
        <el-card shadow="never" style="margin-bottom:16px;">
            <template #header>② 参数试配</template>

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
                            <el-input v-model="params.title" placeholder="标题" style="width:33%;margin-right:8px;" />
                            <el-input v-model="params.xlabel" placeholder="X 轴" style="width:30%;margin-right:8px;" />
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
                            <el-input-number v-model="params.x_major" :min="0" :step="0.5" placeholder="X 步长" /> /
                            <el-input-number v-model="params.y_major" :min="0" :step="0.5" placeholder="Y 步长" />
                            <el-checkbox v-model="params.show_minor_grid" style="margin-left:12px;">次网格</el-checkbox>
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

                        <el-form-item label="下方指标框">
                            <el-switch v-model="params.show_indicator" />
                        </el-form-item>
                    </el-form>
                </el-col>

                <el-col :span="12">
                    <div
                        style="border:1px dashed #ddd; border-radius:8px; padding:12px; min-height:340px; display:flex; flex-direction:column;">
                        <div style="margin-bottom:8px;">
                            <el-button type="primary" size="small" @click="doPreview">刷新预览</el-button>
                            <el-tag type="info" size="small" style="margin-left:8px;">预览 DPI: 120</el-tag>
                        </div>
                        <div style="flex:1; display:flex; align-items:center; justify-content:center;">
                            <img v-if="previewUrl" :src="previewUrl" alt="预览图"
                                style="max-width:100%; max-height:100%; border-radius:6px;" />
                            <div v-else style="color:#999;">暂无预览，请点击“刷新预览”</div>
                        </div>
                    </div>
                </el-col>
            </el-row>
        </el-card>

        <!-- Step 3: 导出 -->
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

            <div v-if="downloadLinks.length || zipLink" style="margin-top:8px;">
                <p style="margin:0 0 6px;">下载链接：</p>
                <ul style="margin:0 0 8px 16px;">
                    <li v-for="(u, idx) in downloadLinks" :key="idx"><a :href="API.defaults.baseURL + u"
                            target="_blank">{{
                                u.split('=').pop() }}</a></li>
                </ul>
                <p v-if="zipLink">打包下载：<a :href="API.defaults.baseURL + zipLink" target="_blank">{{
                    zipLink.split('=').pop()
                        }}</a></p>
            </div>
        </el-card>
    </div>
</template>

<style>
body {
    background: #fafafa;
}
</style>
