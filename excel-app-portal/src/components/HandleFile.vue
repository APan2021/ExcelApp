<template>
  <div class="upload-container">
    <input type="file" id="fileInput" @change="handleFileChange" ref="fileInput" class="file-input" style="display:none;" />
    <label for="fileInput" class="file-input">
      选择Excel文件
    </label>
    <span id="fileName">未选择任何文件</span>
    <button @click="uploadFile" :disabled="isLoading" class="upload-btn">
      上传
    </button>
    <button @click="handleUniqueFile" :disabled="isLoading" class="upload-btn">
      输出单次会诊的病人记录
    </button>

    <button @click="handleMergeData" :disabled="isLoading" class="upload-btn">
      合并多条数据
    </button>

    <div v-if="isLoading" class="loader"></div>

    <!-- 模态框，用于显示处理完成的消息 -->
    <div v-if="showSuccessModal" class="modal">
      <div class="modal-content">
        <p>{{ successMessage }}</p>
        <button @click="closeModal">处理完成</button>
      </div>
    </div>
  </div>
</template>

<script setup>
import {ref} from 'vue';

// 文件选择的引用
const fileInput = ref(null);
// 加载动画的状态
const isLoading = ref(false);
// 模态框的状态和消息
const showSuccessModal = ref(false);
const successMessage = ref('');

const backendURL = process.env.VUE_APP_BACKEND_URL || 'http://localhost:9000';

// 文件大小限制
const MAX_FILE_SIZE = 100 * 1024 * 1024; // 100MB

function handleFileChange(event) {
  const file = event.target.files[0];
  if (file && file.size > MAX_FILE_SIZE) {
    alert(`File size should not exceed ${MAX_FILE_SIZE / 1024 / 1024} MB.`);
    fileInput.value.value = ''; // 清空文件输入
    return;
  }

  const files = event.target.files;
  const fileNameDisplay = document.getElementById('fileName');
  if (files.length > 0) {
    fileNameDisplay.textContent = files[0].name; // 显示文件名
  } else {
    fileNameDisplay.textContent = '未选择任何文件'; // 没有文件时清空显示
  }
}

async function uploadFile() {
  if (!fileInput.value || !fileInput.value.files.length) {
    alert('请先选择一个文件。');
    return;
  }

  const file = fileInput.value.files[0];
  const formData = new FormData();
  formData.append('file', file);

  isLoading.value = true; // 开始加载动画
  try {
    const response = await fetch(`${backendURL}/upload`, {
      method: 'POST',
      body: formData,
    });

    if (!response.ok) {
      throw new Error('上传文件失败。');
    }

    // 上传成功的逻辑
    await response.text();

    successMessage.value = '文件上传成功！'; // 自定义提示信息
    showSuccessModal.value = true; // 打开模态框
  } catch (error) {
    successMessage.value = '上传失败：' + error.message;
    showSuccessModal.value = true;
  } finally {
    isLoading.value = false; // 停止加载动画
  }
}

async function handleUniqueFile() {
  isLoading.value = true; // 开始加载动画
  try {
    const response = await fetch(`${backendURL}/unique`, {
      method: 'GET'
    });

    if (!response.ok) {
      throw new Error('整理单次会诊的病人记录失败');
    }

    // 成功处理后显示模态框
    successMessage.value = '病人记录整理成功！';
    showSuccessModal.value = true;
  } catch (error) {
    successMessage.value = '请求失败：' + error.message;
    showSuccessModal.value = true;
  } finally {
    isLoading.value = false; // 停止加载动画
  }
}

// 处理合并数据的请求
async function handleMergeData() {
  isLoading.value = true;
  try {
    const response = await fetch(`${backendURL}/merge`, {
      method: 'GET',
    });

    if (!response.ok) {
      throw new Error('合并数据失败');
    }

    successMessage.value = '数据合并成功！';
    showSuccessModal.value = true;
  } catch (error) {
    successMessage.value = '合并失败：' + error.message;
    showSuccessModal.value = true;
  } finally {
    isLoading.value = false;
  }
}

// 关闭模态框
function closeModal() {
  showSuccessModal.value = false;
}
</script>

<style scoped>
.upload-container {
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 10px;
  position: absolute;
  z-index: 2;
  left: 50%;
  transform: translate(-50%, -50%);
}

.file-input, .upload-btn {
  background-color: rgba(255,255,255,0.2);
  border: 1px solid rgba(255,255,255,0.4);
  color: #fff;
  padding: 10px 15px;
  border-radius: 3px;
  outline: none;
  transition: background-color 0.25s, border 0.25s;
  cursor: pointer;
  width: 100%;
  text-align: center;
  font-size: 20px;
}

.file-input:hover, .upload-btn:hover {
  background-color: rgba(255,255,255,0.4);
}

.upload-btn:disabled {
  background-color: rgba(255,255,255,0.5);
  color: rgba(255, 255, 255, 0.7);
  cursor: not-allowed;
}

.loader {
  border: 5px solid #f3f3f3;
  border-top: 5px solid #3498db;
  border-radius: 50%;
  width: 50px;
  height: 50px;
  animation: spin 2s linear infinite;
}

@keyframes spin {
  0% { transform: rotate(0deg); }
  100% { transform: rotate(360deg); }
}

/* 模态框样式 */
.modal {
  position: fixed;
  top: 50%;
  left: 50%;
  transform: translate(-50%, -50%);
  z-index: 1000;
  background-color: white;
  padding: 20px;
  box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2);
  border-radius: 8px;
}

.modal-content {
  text-align: center;
}

.modal button {
  margin-top: 10px;
  padding: 5px 10px;
  background-color: #3498db;
  color: white;
  border: none;
  border-radius: 4px;
  cursor: pointer;
}

.modal button:hover {
  background-color: #2980b9;
}
</style>
