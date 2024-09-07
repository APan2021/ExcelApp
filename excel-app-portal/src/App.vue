<template>
  <div class="container">
    <file-upload></file-upload>
  </div>
  <ul class="bg-squares">
    <li></li>
    <li></li>
    <li></li>
    <li></li>
    <li></li>
    <li></li>
    <li></li>
    <li></li>
    <li></li>
    <li></li>
  </ul>
</template>

<script setup>
import FileUpload from './components/HandleFile.vue';
</script>

<style>
* {
  /* 初始化 */
  margin: 0;
  padding: 0;
  box-sizing: border-box;
}
body {
  height: 100vh;
  /* 弹性布局 居中显示 */
  display: flex;
  justify-content: center;
  align-items: center;
  /* 渐变背景 */
  background: linear-gradient(to top left,#5f7bad,#577dad,#416faf);
  /* 溢出隐藏 */
  overflow: hidden;
}
.container {
  text-align: center;
  color: #fff;
}
.container h1 {
  font-size: 40px;
  font-weight: 100;
  letter-spacing: 2px;
  margin-bottom: 15px;
  /* 过渡效果 */
  transition: 1s ease-in-out;
}



/* 背景方块 */
.bg-squares {
  list-style: none;
  position: absolute;
  top: 0;
  left: 0;
  width: 100%;
  height: 100%;
  z-index: 1;
}
.bg-squares li {
  width: 40px;
  height: 40px;
  background-color: rgba(255,255,255,0.15);
  position: absolute;
  bottom: -160px;
  /* 执行动画：动画名 时长 线性 无限次播放 */
  animation: square 20s linear infinite;
}
/* 为每一个方块设置不同的位置、大小、动画延迟时间、动画时长、背景色 */
.bg-squares li:nth-child(1) {
  left: 10%;
}
.bg-squares li:nth-child(2) {
  left: 20%;
  width: 80px;
  height: 80px;
  /* 动画延迟时间 */
  animation-delay: 2s;
  /* 动画时长 */
  animation-duration: 17s;
}
.bg-squares li:nth-child(3) {
  left: 25%;
  animation-delay: 4s;
}
.bg-squares li:nth-child(4) {
  left: 40%;
  width: 60px;
  height: 60px;
  background-color: rgba(255,255,255,0.25);
  animation-duration: 22s;
}
.bg-squares li:nth-child(5) {
  left: 70%;
}
.bg-squares li:nth-child(6) {
  left: 80%;
  width: 120px;
  height: 120px;
  background-color: rgba(255,255,255,0.2);
  animation-delay: 3s;
}
.bg-squares li:nth-child(7) {
  left: 32%;
  width: 160px;
  height: 160px;
  animation-delay: 7s;
}
.bg-squares li:nth-child(8) {
  left: 55%;
  width: 20px;
  height: 20px;
  animation-delay: 15s;
  animation-duration: 40s;
}
.bg-squares li:nth-child(9) {
  left: 25%;
  width: 10px;
  height: 10px;
  background-color: rgba(255,255,255,0.3);
  animation-delay: 2s;
  animation-duration: 40s;
}
.bg-squares li:nth-child(10) {
  left: 90%;
  width: 160px;
  height: 160px;
  animation-delay: 11s;
}
.container.success h1{
  transform: translateY(75px);
}

/* 定义动画 */
@keyframes square {
  0%{
    transform: translateY(0);
  }
  100%{
    transform: translateY(-120vh) rotate(600deg);
  }
}
</style>

<style>
html, body {
  margin: 0;
  padding: 0;
  overflow: hidden; /* 防止滚动条影响背景填充 */
  height: 100%;
}
</style>
