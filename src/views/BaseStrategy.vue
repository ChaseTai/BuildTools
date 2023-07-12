<template>
  <div class="pageContent">
    <el-button class="my_file" :icon="Upload" type="primary">
      <input type="file" class="my_input" id="upload" @change="handleChange">
      导入EXCEL
    </el-button>
    <el-table :data="tableData" style="width: 100%">
      <el-table-column v-for="(item, index) in columns" :prop="item.prop" :label="item.label" :width="item.width" />
    </el-table>
    <div class="output" v-html="moduleContent"></div>
  </div>
</template>

<script setup>
  import {Upload} from "@element-plus/icons-vue";
  import {ElMessage} from "element-plus";
  import * as ExcelUtil from "@/utils/handleExcel.js";

  const excelFiles = ref(null)
  const moduleContent = ref(null)
  const columns = [
    {label: '策略名', prop: 'strategyName'},
    {label: '描述', prop: 'description'},
    {label: '描述原文', prop: 'descriptionOrigin'},
    {label: '源安全区域', prop: 'sourceSafeZone'},
    {label: '目的安全区域', prop: 'destinationSafeZone'},
    {label: '源地址', prop: 'sourceAddress'},
    {label: '目的地址', prop: 'destinationAddress'},
    {label: '协议', prop: 'protocol'},
    {label: '服务端口(必填)', prop: 'servicePorts'},
    {label: '服务类型(必填)', prop: 'serviceType'},
  ]
  const tableData = ref([])

  onMounted(()=>{
    document.getElementById('upload').value = '';
  })

  // 导入excel
  const handleChange = (e) => {
    const files = e.target.files;
    excelFiles.value = e.target.files;
    if (!files.length) {
      return ;
    } else if (!/\.(xls|xlsx)$/.test(files[0].name.toLowerCase())) {
      ElMessage.warning('上传格式不正确，请上传xls或者xlsx格式');
    }
    getExcelData();
  }
  // 切换tab获取对应excel内容
  const getExcelData = () => {
    ExcelUtil.getExcelData({
      tag: 'Weekly',
      excelFiles: excelFiles.value[0],
      callback: renderExcel
    });
  }
  // 将解析后的excel数据渲染到表格
  const renderExcel = (data) => {
    console.log(data)
  }
  const getContent = () => {
    console.log(moduleContent.value)
  }
</script>

<style lang="scss" scoped>
.module-area{
  margin-top: 10px;
  .module-content{
    .content{
      width: 400px;
      height: 340px;
    }
  }
}
</style>