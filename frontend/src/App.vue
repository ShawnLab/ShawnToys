<script lang="ts" setup>
import Home from './components/Home.vue'
import Logo from './components/Logo.vue'
</script>

<template>
  <Logo id="logo"/>
  <Home/>
</template>

<style>
#logo {
  display: block;
  width: 50%;
  height: 30%;
  margin: auto;
  padding: 10% 0 0;
  background-position: center;
  background-repeat: no-repeat;
  background-size: 100% 100%;
  background-origin: content-box;
}
</style>
