<template>
  <div class="leftNav">
    <el-menu
        :default-active="route.path"
        active-text-color="#ffd04b"
        background-color="#545c64"
        class="el-menu-vertical-demo"
        text-color="#fff"
        router
    >
      <el-menu-item index="/home">
        <el-icon><location /></el-icon>
        <span>首页</span>
      </el-menu-item>
      <el-menu-item index="/baseStrategy">
        <el-icon><icon-menu /></el-icon>
        <span>办公室基础策略</span>
      </el-menu-item>
      <el-menu-item index="/baseStrategy">
        <el-icon><icon-menu /></el-icon>
        <span>AP-SN</span>
      </el-menu-item>
    </el-menu>
  </div>
</template>

<script setup>
  import { Menu as IconMenu, Location, } from '@element-plus/icons-vue'
  import {useRoute} from "vue-router";

  const route = useRoute();
</script>

<style lang="scss" scoped>
.leftNav{
  height: 100%;
  .el-menu{
    height: 100%;
  }
}
</style>