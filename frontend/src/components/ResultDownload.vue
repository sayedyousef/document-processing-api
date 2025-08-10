<template>
  <div class="max-w-xl mx-auto mt-6 p-6 bg-white rounded-lg shadow-md">
    <h3 class="text-lg font-semibold mb-4">Results Ready</h3>
    
    <div class="space-y-3">
      <div 
        v-for="(result, index) in results" 
        :key="index"
        class="flex justify-between items-center p-3 bg-gray-50 rounded-lg"
      >
        <div class="flex items-center space-x-3">
          <svg class="h-5 w-5 text-green-500" fill="currentColor" viewBox="0 0 20 20">
            <path fill-rule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clip-rule="evenodd" />
          </svg>
          <div>
            <p class="font-medium">{{ result.filename }}</p>
            <p class="text-sm text-gray-600">
              {{ formatFileSize(result.size) }}
              <span v-if="result.type === 'application/zip'" class="ml-2 text-blue-600">
                (Multiple files packaged)
              </span>
            </p>
          </div>
        </div>
        
        <div class="flex items-center space-x-2">
          <span v-if="result.error" class="text-xs text-red-500">{{ result.error }}</span>
          <button 
            v-else
            @click="downloadFile(index)"
            class="px-3 py-1 bg-blue-500 text-white text-sm rounded hover:bg-blue-600 transition-colors"
          >
            {{ result.type === 'application/zip' ? '📦 Download ZIP' : 'Download' }}
          </button>
        </div>
      </div>
    </div>
    
    <button 
      @click="downloadAll"
      v-if="results.length > 1 && !hasErrors"
      class="mt-4 w-full px-4 py-2 bg-green-500 text-white rounded hover:bg-green-600 transition-colors font-semibold"
    >
      Download All as ZIP
    </button>
    
    <button 
      @click="$emit('reset')"
      class="mt-3 w-full px-4 py-2 bg-gray-200 text-gray-700 rounded hover:bg-gray-300 transition-colors"
    >
      Process More Documents
    </button>
  </div>
</template>

<script>
import { computed } from 'vue'

export default {
  props: ['results', 'jobId'],
  emits: ['reset'],
  setup(props) {
    const hasErrors = computed(() => {
      return props.results.some(r => r.error)
    })
    
    // FORMAT FILE SIZE FUNCTION
    const formatFileSize = (bytes) => {
      console.log('[ResultDownload] Formatting file size:', bytes)
      if (!bytes || bytes === 0) return '0 Bytes'
      const k = 1024
      const sizes = ['Bytes', 'KB', 'MB', 'GB']
      const i = Math.floor(Math.log(bytes) / Math.log(k))
      return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i]
    }
    
    const downloadFile = (index) => {
      const url = `http://localhost:8000/api/download/${props.jobId}/${index}`
      console.log('[ResultDownload] Downloading file from:', url)
      console.log('[ResultDownload] File info:', props.results[index])
      
      // Create download link
      const link = document.createElement('a')
      link.href = url
      link.download = props.results[index].filename || `download_${index}`
      document.body.appendChild(link)
      link.click()
      document.body.removeChild(link)
    }
    
    const downloadAll = () => {
      const url = `http://localhost:8000/api/download/${props.jobId}`
      console.log('[ResultDownload] Downloading all as ZIP from:', url)
      
      const link = document.createElement('a')
      link.href = url
      link.download = `results_${props.jobId}.zip`
      document.body.appendChild(link)
      link.click()
      document.body.removeChild(link)
    }
    
    console.log('[ResultDownload] Component setup with results:', props.results)
    
    return {
      downloadFile,
      downloadAll,
      hasErrors,
      formatFileSize  // ← THIS WAS MISSING! Must be returned
    }
  }
}
</script>