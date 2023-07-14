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
        <el-button @click="downloadTemplate" type="primary" :icon="Download" link size="small" v-show="isConfirm">下载表格</el-button>
      </el-form-item>
    </el-form>
    <el-table :data="tableData" :table-layout="'auto'" style="width: 100%">
      <el-table-column v-for="(item, index) in columns" :prop="item.prop" :label="item.label" :width="item.width" :key="index">
        <template v-if="item.prop==='description'" #default="scope">
          <p>描述：{{scope.row.description}}</p>
          <p>原文：{{scope.row.descriptionOrigin}}</p>
        </template>
        <template v-else-if="item.prop==='sourceLogical'" #default="scope">
          <p>源逻辑实体：{{scope.row.sourceLogical}}</p>
          <p>目的逻辑实体：{{scope.row.destinationLogical}}</p>
        </template>
        <template v-else-if="item.prop==='sourceIP'" #default="scope">
          <p>源地址</p>
          <p>{{scope.row.sourceIP}}</p>
          <p>目的地址</p>
          <p v-for="sub in scope.row.destinationIP">{{sub}}</p>
        </template>
        <template v-else-if="item.prop==='servicePorts'" #default="scope">
          <p v-for="sub in scope.row.servicePorts">{{sub}}</p>
        </template>
        <template v-else-if="item.prop==='serviceType'" #default="scope">
          {{scope.row.serviceType.length > 1 ? scope.row.serviceType.join('/') : scope.row.serviceType.toString()}}
        </template>
        <template v-else-if="item.prop==='orderGroup'" #default="scope">
          <div v-html="scope.row.orderGroup"></div>
        </template>
      </el-table-column>
    </el-table>
  </div>
</template>

<script setup>
  import {Upload, Download} from "@element-plus/icons-vue";
  import {ElMessage} from "element-plus";
  import * as ExcelUtil from "@/utils/handleExcel.js";
  import {exportJsonToExcel} from "@/utils/Export2Excel.js";

  const queryFormRef = ref()
  const queryForm = ref({
    wifiZone: '福州分行'  // 定义区域为WiFi
  })
  const isConfirm = ref(false)
  const excelFiles = ref(null)
  const columns = ref([
    {label: '策略名', prop: 'strategyName', width: '140px'},
    {label: '描述', prop: 'description', width: '300px'},
    // {label: '描述原文', prop: 'descriptionOrigin'},
    {label: '逻辑实体', prop: 'sourceLogical', width: '200px'},
    // {label: '目的逻辑实体', prop: 'destinationLogical'},
    {label: '地址', prop: 'sourceIP', width: '200px'},
    // {label: '目的地址', prop: 'destinationIP'},
    {label: '命令组合', prop: 'orderGroup'},
    {label: '服务端口', prop: 'servicePorts', width: '200px'},
    {label: '服务类型', prop: 'serviceType', width: '100px'},
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
      if (item.destinationIP) {
        item.destinationIP = item.destinationIP.replace(/\//g, ' ');
        if (item.destinationIP.indexOf('\r\n') > -1) {
          item.destinationIP = item.destinationIP.split('\r\n').filter(i=>i);
        } else {
          item.destinationIP = [item.destinationIP];
        }
      }

      if (String(item.sourceIP).indexOf('/') > -1) {item.sourceIP = item.sourceIP.replace('/', ' ')}

      let tcpType = false, udpType = false;
      let newArr = String(item.servicePorts).replace(/[\r\n,、]/g, ' ').split(' ').filter(i=>i);

      item.servicePorts = newArr.map(sub => {
        if (sub.indexOf('tcp：') > -1) { sub = sub.replace('tcp：', ''); tcpType = true; }
        else { tcpType = false; }
        if (sub.indexOf('udp：') > -1) { sub = sub.replace('udp：', ''); udpType = true; }
        else { udpType = false; }
        if (sub.indexOf('-') > -1) { sub = sub.replace('-', ' to ')}
        return sub;
      });

      if (!item.serviceType) {
          if (tcpType && udpType) {
              item.serviceType = 'tcp/udp'
          } else {
              if (tcpType) { item.serviceType = 'tcp' }
              else if (udpType) { item.serviceType = 'udp' }
          }
      }

      item.serviceType = item.serviceType ? item.serviceType.toUpperCase() : 'TCP';
      if (item.serviceType && String(item.serviceType).indexOf('/') > -1) {
        item.serviceType = item.serviceType.split('/');
      } else {
        if (item.serviceType.indexOf('\r\n') > -1) {
          item.serviceType = item.serviceType.split('\r\n');
        } else {
          item.serviceType = [item.serviceType];
        }
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

      item.orderGroup = `rule name ${item.strategyName}<br/>description ${item.description}<br/>${item.sourceSafeZone}<br/>${item.destinationSafeZone}<br/>source-address ${item.sourceIP}<br/>${destinationIPs}${servicePorts}policy logging<br/>session logging<br/>service icmp<br/>action permit<br/>#`;
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
    if (tableDataR.length > 0) {
      tableDataR.forEach(item=>{
        let destinationIPs = item.destinationIP.map(sub=>sub+'\r');
        item.description = '描述：'+item.description+'\r'+'原文：'+item.descriptionOrigin+'\r';
        item.sourceLogical = '源逻辑实体：'+item.sourceLogical+'\r'+'目的逻辑实体：'+item.destinationLogical+'\r';
        item.sourceIP = '源地址：'+item.sourceIP+'\r'+'目的地址：'+destinationIPs;
        item.servicePorts = item.servicePorts.join('\r');
        item.serviceType = item.serviceType.length > 1 ? item.serviceType.join('/') : item.serviceType.toString();
        item.orderGroup = item.orderGroup.replaceAll('<br/>', '\r');
      })
    }
    console.log(tableDataR)

    const data = formatJson(filterVal, tableDataR);
    // 可以根据项目需要传入 bookType （文件后缀名）
    exportJsonToExcel({
      header: tHeader,
      data,
      filename: '办公室基础策略'
    });
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