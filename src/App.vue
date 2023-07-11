<script setup>
import LeftNav from "@/components/LeftNav.vue";
</script>

<template>
  <div class="common-layout">
    <el-container>
      <el-aside width="200px"><LeftNav /></el-aside>
      <el-container>
        <el-header>Header</el-header>
        <el-main><router-view /></el-main>
      </el-container>
    </el-container>
  </div>
</template>

<style scoped>
</style>
