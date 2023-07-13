<template>
  <div class="pageContent">
    <el-form ref="queryFormRef" :model="queryForm" inline label-suffix=" :">
      <el-form-item label="定义WIFI区域">
        <el-input v-model="queryForm.wifiZone" placeholder="请输入" :disabled="isConfirm" />
      </el-form-item>
      <el-form-item>
        <el-button type="primary" @click="actionClick('confirm')">确认</el-button>
        <el-button @click="actionClick('reset')">重置</el-button>
        <el-button class="my_file" :icon="Upload" type="primary" v-show="isConfirm">
          <input type="file" class="my_input" id="upload" @change="handleChange">
          导入EXCEL
        </el-button>
        <el-button @click="downloadTemplate" type="primary" :icon="Download" link size="small" v-show="isConfirm">下载导入模板</el-button>
      </el-form-item>
    </el-form>
    <el-table :data="tableData" :table-layout="'auto'" style="width: 100%">
      <el-table-column v-for="(item, index) in columns" :prop="item.prop" :label="item.label" :width="item.width" :key="index">
        <template v-if="item.prop==='destinationIP'" #default="scope">
          <p v-for="sub in scope.row.destinationIP">{{sub}}</p>
        </template>
        <template v-else-if="item.prop==='servicePorts'" #default="scope">
          <p v-for="sub in scope.row.servicePorts">{{sub}}</p>
        </template>
        <template v-else-if="item.prop==='orderGroup'" #default="scope">
          <div v-html="scope.row.orderGroup"></div>
        </template>
      </el-table-column>
    </el-table>
    <div class="output" v-html="moduleContent"></div>
  </div>
</template>

<script setup>
  import {Upload, Download} from "@element-plus/icons-vue";
  import {ElMessage} from "element-plus";
  import * as ExcelUtil from "@/utils/handleExcel.js";
  import {exportJsonToExcel} from "@/utils/Export2Excel.js";

  const queryFormRef = ref()
  const queryForm = ref({
    wifiZone: null  // 定义区域为WiFi
  })
  const isConfirm = ref(false)
  const excelFiles = ref(null)
  const moduleContent = ref(null)
  const columns = ref([
    {label: '策略名', prop: 'strategyName'},
    {label: '描述', prop: 'description'},
    {label: '描述原文', prop: 'descriptionOrigin'},
    {label: '源逻辑实体', prop: 'sourceLogical'},
    {label: '目的逻辑实体', prop: 'destinationLogical'},
    // {label: '源安全区域', prop: 'sourceSafeZone'},
    // {label: '目的安全区域', prop: 'destinationSafeZone'},
    {label: '源地址', prop: 'sourceIP'},
    {label: '目的地址', prop: 'destinationIP'},
    // {label: '协议', prop: 'protocol'},
    {label: '命令组合', prop: 'orderGroup'},
    {label: '服务端口(必填)', prop: 'servicePorts'},
    {label: '服务类型(必填)', prop: 'serviceType'},
  ])
  const tableData = ref([])

  onMounted(()=>{
    document.getElementById('upload').value = '';
  })

  // 事件点击
  const actionClick = (val) => {
    isConfirm.value = val === 'confirm';
  }
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
    data.forEach(item=>{
      if (item.destinationIP && item.destinationIP.indexOf('\r\n') > -1) {
        item.destinationIP = item.destinationIP.split('\r\n').filter(i=>i);
      } else {
        item.destinationIP = [item.destinationIP];
      }
      if (item.servicePorts && String(item.servicePorts).indexOf('\r\n') > -1) {
        item.servicePorts = item.servicePorts.split('\r\n').filter(i=>i);
      } else {
        item.servicePorts = [item.servicePorts];
      }
      item.serviceType = item.serviceType.toUpperCase();
      if (item.serviceType && String(item.serviceType).indexOf('/') > -1) {
        item.serviceType = item.serviceType.split('/');
      } else {
        item.serviceType = [item.serviceType];
      }

      item.sourceSafeZone = item.sourceLogical === queryForm.value.wifiZone ? 'source-zone WIFI' : 'source-zone OA';
      item.destinationSafeZone = item.destinationLogical === queryForm.value.wifiZone ? 'destination-zone WIFI' : 'destination-zone OA';

      let destinationIPs = '';
      for (let i = 0; i < item.destinationIP.length; i++) {
        destinationIPs += `destination-address ${item.destinationIP[i]}<br/>`
      }
      let servicePorts = '';
      for (let i = 0; i < item.serviceType.length; i++) {
        for (let j = 0; j < item.servicePorts.length; j++) {
          servicePorts += `service protocol ${item.serviceType[i] === 'HTTP' ? 'tcp' : item.serviceType[i].toLowerCase()} destination-port ${item.servicePorts[j]}<br/>`
        }
      }

      item.orderGroup = `rule name ${item.strategyName}<br/>
        description ${item.description}<br/>
        ${item.sourceSafeZone}<br/>
        ${item.destinationSafeZone}<br/>
        source-address ${item.sourceIP}<br/>
        ${destinationIPs}
        ${servicePorts}
        policy logging<br/>
        session logging<br/>
        service icmp<br/>
        action permit<br/>
        #`
    })
    console.log(data)
    tableData.value = data;
  }
  // 获取每个数据
  const formatJson = (filterVal, jsonData) => {
    return jsonData.map((v) =>
        filterVal.map((j) => {
          // 此判断解决数据多层嵌套
          if (j.indexOf('.') > 0) {
            const arr = j.split('.');
            let resData = v;
            for (var i = 0; i < arr.length; i++) {
              resData = resData[arr[i]];
            }
            return resData;
          } else {
            return v[j];
          }
        })
    );
  }
  // 下载导入模板
  const downloadTemplate = () => {
    // 对应表格的 label
    let tHeader = columns.value.map(item=>item.label);
    // 对应表格的 prop
    let filterVal = columns.value.map(item=>item.prop);
    let tableDataR = JSON.parse(JSON.stringify(tableData.value));
    tableDataR.forEach(item=>{
      item.servicePorts = item.servicePorts.join('\r');
      item.destinationIP = item.destinationIP.join('\r');
      item.orderGroup = item.orderGroup.replaceAll('<br/>', '\r');
    })
    console.log(tableDataR)

    const data = formatJson(filterVal, tableDataR);
    // 可以根据项目需要传入 bookType （文件后缀名）
    exportJsonToExcel({
      header: tHeader,
      data,
      filename: '办公室基础策略'
    });
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