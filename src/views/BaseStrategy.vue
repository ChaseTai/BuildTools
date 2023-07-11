<template>
  <div>
    <el-button class="my_file" :icon="Upload" type="primary">
      <input type="file" class="my_input" id="upload" @change="handleChange">
      导入EXCEL
    </el-button>
  </div>
</template>

<script setup>
  import {Upload} from "@element-plus/icons-vue";
  import {ElMessage} from "element-plus";
  import * as ExcelUtil from "@/utils/handleExcel.js";

  const excelFiles = ref(null)

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
      // callback: renderExcel
    });
  }
</script>

<style scoped>

</style>