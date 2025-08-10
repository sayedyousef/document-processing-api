<template>
  <div v-if="status" class="max-w-xl mx-auto mt-6 p-6 bg-white rounded-lg shadow-md">
    <h3 class="text-lg font-semibold mb-4">Processing Status</h3>
    
    <div class="space-y-3">
      <!-- Status indicator -->
      <div class="flex items-center space-x-3">
        <div class="relative">
          <div v-if="status.status === 'processing'" class="animate-spin h-5 w-5 border-2 border-blue-500 border-t-transparent rounded-full"></div>
          <svg v-else-if="status.status === 'completed'" class="h-5 w-5 text-green-500" fill="currentColor" viewBox="0 0 20 20">
            <path fill-rule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clip-rule="evenodd" />
          </svg>
          <svg v-else class="h-5 w-5 text-red-500" fill="currentColor" viewBox="0 0 20 20">
            <path fill-rule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zM8.707 7.293a1 1 0 00-1.414 1.414L8.586 10l-1.293 1.293a1 1 0 101.414 1.414L10 11.414l1.293 1.293a1 1 0 001.414-1.414L11.414 10l1.293-1.293a1 1 0 00-1.414-1.414L10 8.586 8.707 7.293z" clip-rule="evenodd" />
          </svg>
        </div>
        <span class="font-medium capitalize">{{ status.status }}</span>
      </div>
      
      <!-- Progress -->
      <div>
        <div class="flex justify-between text-sm text-gray-600 mb-1">
          <span>Progress</span>
          <span>{{ status.progress }}</span>
        </div>
        <div class="w-full bg-gray-200 rounded-full h-2">
          <div 
            class="bg-blue-500 h-2 rounded-full transition-all duration-300"
            :style="{ width: progressPercentage + '%' }"
          ></div>
        </div>
      </div>
      
      <!-- Processor type -->
      <div class="text-sm text-gray-600">
        <span class="font-medium">Processor:</span> 
        {{ status.processor === 'scan_verify' ? 'Document Analysis' : 
           status.processor === 'word_to_html' ? 'HTML Conversion' : 
           status.processor === 'latex_equations' ? 'LaTeX Equation Extraction' :
           status.processor === 'word_complete' ? 'Complete Word to HTML Processing' :
           'Processing' }}
      </div>
    </div> <!-- ADDED: Closes space-y-3 div -->
  </div> <!-- ADDED: Closes v-if="status" div -->
</template>

<script>
import { ref, computed, onMounted, onUnmounted } from 'vue'
import axios from 'axios'

export default {
  props: ['jobId'],
  emits: ['completed'],
  setup(props, { emit }) {
    const status = ref(null)
    const polling = ref(null)
    
    const progressPercentage = computed(() => {
      if (!status.value || !status.value.progress) return 0
      const [completed, total] = status.value.progress.split('/').map(Number)
      return total > 0 ? (completed / total) * 100 : 0
    })
    
    const checkStatus = async () => {
      try {
        const response = await axios.get(`http://localhost:8000/api/status/${props.jobId}`)
        status.value = response.data
        
        if (response.data.status === 'completed') {
          clearInterval(polling.value)
          emit('completed', response.data.results)
        }
      } catch (error) {
        console.error('Status check failed:', error)
      }
    }
    
    onMounted(() => {
      checkStatus()
      polling.value = setInterval(checkStatus, 2000)
    })
    
    onUnmounted(() => {
      if (polling.value) {
        clearInterval(polling.value)
      }
    })
    
    return {
      status,
      progressPercentage
    }
  }
}
</script>