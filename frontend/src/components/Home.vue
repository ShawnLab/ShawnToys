<template>
    <div class="home">
        <h1>Excel Sheet Text Replacer</h1>
        <form @submit.prevent="replaceText">
            <div class="input-box">
                <label for="filePath">Excel文件：</label>
                <input class="input" id="filePath" v-model="excelFilePath" readonly>
                <button class="btn" type="button" @click="selectFile('excel')">选择文件</button>
            </div>
            <br />
            <div class="input-box">
                <label for="rulePath">Rule文件：</label>
                <input class="input" id="rulePath" v-model="ruleFilePath" readonly>
                <button class="btn" type="button" @click="selectFile('rule')">选择文件</button>
            </div>
            <br />
            <button class="sub-btn" type="submit">替换</button>
        </form>
        <div class="result" v-if="error">{{ error }}</div>
        <div class="result" v-if="success">{{ success }}</div>
    </div>
</template>

<script lang="ts" setup>
import { ref } from 'vue';
import { ReplaceTextInSheet, SelectFile } from '../../wailsjs/go/main/App';
import { LogPrint } from '../../wailsjs/runtime';

const excelFilePath = ref()
const ruleFilePath = ref()
const error = ref('')
const success = ref('')

async function selectFile(target: string) {
    try {
        switch (target) {
            case 'excel':

                excelFilePath.value = await SelectFile('Select Excel File', 'Excel Files', '*.xlsx;*.xls');
                break;
            case 'rule':
                ruleFilePath.value = await SelectFile('Select Rule File', 'Rule Files', '*.txt');
                break;

            default:
                break;
        }

    } catch (err) {
        error.value = `Error: ${err}`;
        success.value = '';
    }
}
async function replaceText() {
    if (!excelFilePath.value || !ruleFilePath.value) {
        if (excelFilePath.value) { error.value = 'Rule文件必选'; return }
        else if (ruleFilePath.value) { error.value = 'Excel文件必选'; return }
        error.value = 'Excel文件和Rule文件必选'; return
    }
    try {
        const result = await ReplaceTextInSheet(excelFilePath.value, ruleFilePath.value);
        success.value = result;
        error.value = '';
    } catch (err) {
        LogPrint(typeof err)
        error.value = `Error: ${err}`;
        success.value = '';
    }
}

</script>

<style scoped>
.result {
    height: 20px;
    line-height: 20px;
    margin: 1.5rem auto;
}

.input-box .btn {
    width: 80px;
    height: 30px;
    line-height: 30px;
    border-radius: 3px;
    border: none;
    margin: 0 0 0 20px;
    padding: 0 8px;
    cursor: pointer;
}

.input-box .btn:hover {
    background-image: linear-gradient(to top, #cfd9df 0%, #e2ebf0 100%);
    color: #333333;
}

.input-box .input {
    border: none;
    border-radius: 3px;
    outline: none;
    height: 30px;
    line-height: 30px;
    padding: 0 10px;
    background-color: rgba(240, 240, 240, 1);
    -webkit-font-smoothing: antialiased;
}

.input-box .input:hover {
    border: none;
    background-color: rgba(255, 255, 255, 1);
}

.input-box .input:focus {
    border: none;
    background-color: rgba(255, 255, 255, 1);
}

.sub-btn {
    width: 80px;
    height: 30px;
    line-height: 30px;
    border-radius: 3px;
    border: none;
    margin: 0 0 0 20px;
    padding: 0 8px;
    cursor: pointer;
    background-color: #1E6FFF;
    color: white;
}

.sub-btn:hover {
    background-image: linear-gradient(to top, #cfd9df 0%, #e2ebf0 100%);
    color: #333333;
}
</style>