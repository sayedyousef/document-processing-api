<template>
    <div class="max-w-xl mx-auto p-6">
      <div 
        @drop="handleDrop"
        @dragover.prevent
        @dragenter.prevent
        class="border-2 border-dashed border-gray-300 rounded-lg p-8 text-center hover:border-blue-500 transition-colors cursor-pointer"
        :class="{ 'border-blue-500 bg-blue-50': isDragging }"
      >
        <svg class="mx-auto h-12 w-12 text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" 
                d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
        </svg>
        <p class="mt-2 text-sm text-gray-600">Drop Word documents here or click to select</p>
        <input 
          type="file" 
          multiple 
          accept=".docx,.doc"
          @change="handleFileSelect" 
          class="hidden" 
          ref="fileInput"
        >
        <button 
          @click="$refs.fileInput.click()" 
          class="mt-4 px-4 py-2 bg-blue-500 text-white rounded hover:bg-blue-600 transition-colors"
        >
          Select Files
        </button>
      </div>
      
      <!-- File list -->
      <div v-if="files.length" class="mt-6">
        <h3 class="text-lg font-semibold mb-3">Selected Files ({{ files.length }})</h3>
        <div class="space-y-2">
          <div 
            v-for="(file, index) in files" 
            :key="index"
            class="flex justify-between items-center p-3 bg-white rounded-lg shadow-sm border border-gray-200"
          >
            <div class="flex items-center space-x-3">
              <svg class="h-5 w-5 text-blue-500" fill="currentColor" viewBox="0 0 20 20">
                <path d="M9 2a2 2 0 00-2 2v8a2 2 0 002 2h6a2 2 0 002-2V6.414A2 2 0 0016.414 5L14 2.586A2 2 0 0012.586 2H9z" />
              </svg>
              <span class="text-sm font-medium">{{ file.name }}</span>
            </div>
            <div class="flex items-center space-x-3">
              <span class="text-xs text-gray-500">{{ formatSize(file.size) }}</span>
              <button 
                @click="removeFile(index)"
                class="text-red-500 hover:text-red-700"
              >
                <svg class="h-5 w-5" fill="currentColor" viewBox="0 0 20 20">
                  <path fill-rule="evenodd" 
                        d="M4.293 4.293a1 1 0 011.414 0L10 8.586l4.293-4.293a1 1 0 111.414 1.414L11.414 10l4.293 4.293a1 1 0 01-1.414 1.414L10 11.414l-4.293 4.293a1 1 0 01-1.414-1.414L8.586 10 4.293 5.707a1 1 0 010-1.414z" 
                        clip-rule="evenodd" />
                </svg>
              </button>
            </div>
          </div>
        </div>
        
        <button 
          @click="$emit('upload')"
          class="mt-4 w-full px-4 py-2 bg-green-500 text-white rounded hover:bg-green-600 transition-colors font-semibold"
          :disabled="files.length === 0"
        >
          Process {{ files.length }} Document{{ files.length !== 1 ? 's' : '' }}
        </button>
      </div>
    </div>
  </template>
  
  <script>
  import { ref } from 'vue'
  
  export default {
    props: ['files'],
    emits: ['files-selected', 'upload'],
    setup(props, { emit }) {
      const isDragging = ref(false)
      
      const handleDrop = (e) => {
        e.preventDefault()
        isDragging.value = false
        const droppedFiles = Array.from(e.dataTransfer.files).filter(
          file => file.name.endsWith('.docx') || file.name.endsWith('.doc')
        )
        emit('files-selected', droppedFiles)
      }
      
      const handleFileSelect = (e) => {
        const selectedFiles = Array.from(e.target.files)
        emit('files-selected', selectedFiles)
      }
      
      const removeFile = (index) => {
        const newFiles = [...props.files]
        newFiles.splice(index, 1)
        emit('files-selected', newFiles)
      }
      
      const formatSize = (bytes) => {
        if (bytes === 0) return '0 Bytes'
        const k = 1024
        const sizes = ['Bytes', 'KB', 'MB', 'GB']
        const i = Math.floor(Math.log(bytes) / Math.log(k))
        return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i]
      }
      
      return {
        isDragging,
        handleDrop,
        handleFileSelect,
        removeFile,
        formatSize
      }
    }
  }
  </script>