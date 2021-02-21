<template>
    <section class="section-hero section-shaped my-0">
        <div class="shape shape-style-1 shape-primary">
            <span class="span-150"></span>
            <span class="span-50"></span>
            <span class="span-50"></span>
            <span class="span-75"></span>
            <span class="span-100"></span>
            <span class="span-75"></span>
            <span class="span-50"></span>
            <span class="span-100"></span>
            <span class="span-50"></span>
            <span class="span-100"></span>
        </div>
        <div class="container shape-container d-flex align-items-center">
            <div class="col px-0">
                <div class="row justify-content-center align-items-center">
                    <div class="text-center">
                        <p class="lead text-white mt-4 mb-5">Convert your annoying attendance excel to beautiful copasable excel.</p>
                        <base-alert type="danger" v-if="showError">
                            <span class="alert-inner--icon"><i class="fa fa-exclamation-triangle"></i></span>
                            <span class="alert-inner--text"><strong>Error!</strong> Sheet format not match!</span>
                        </base-alert>
                        <vue-dropzone 
                            id="dropzone"
                            @vdropzone-file-added="vfileAdded"
                            @vdropzone-removed-file="vremoved"
                            class="mb-5" 
                            :options="dropzoneOptions" />

                        <base-button 
                            v-for="(sheet, index) in sheets"
                            v-bind:key="index"
                            @click="convertAttendance(sheet)"
                            class="mb-3 mb-sm-0"
                            type="white"
                            icon="ni ni-cloud-download-95">
                            {{ sheet }}
                        </base-button>
                    </div>
                </div>
                <div class="row align-items-center justify-content-around stars-and-coded">
                    <div class="col-sm-4">
                        <span class="text-white alpha-7 ml-3">Come to my</span>
                        <a href="https://github.com/shancheas" target="_blank" title="Support us on Github">
                            <img src="img/brand/github-white-slim.png" style="height: 22px; margin-top: -3px">
                        </a>
                    </div>
                    <div class="col-sm-4 mt-4 mt-sm-0 text-right">
                        <span class=" text-white alpha-7 d-none d-sm-inline-block">Created with <i class="fa fa-heart text-danger"></i> by shancheas.</span>
                    </div>
                </div>
            </div>
        </div>
    </section>
</template>
<script>
import vue2Dropzone from 'vue2-dropzone'
import XLSX from 'xlsx'
import moment from 'moment'
import 'vue2-dropzone/dist/vue2Dropzone.min.css'

const START_NAME_COLUMN = 7
const DAY_ROWS = [
    'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J',
    'K', 'L', 'M', 'O', 'P', 'Q', 'R', 'S', 'T', 'U',
    'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF'
]

export default {
  components: {
    vueDropzone: vue2Dropzone
  },
  data () {
    return {
      sheets: [],
      xlsxFile: null,
      showError: false,
      dropzoneOptions: {
        acceptedFiles: 'application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,xlsx,xls',
        url: 'https://httpbin.org/post',
        thumbnailWidth: 200,
        addRemoveLinks: true,
        dictDefaultMessage: "<i class='fa fa-cloud-upload'></i>UPLOAD ME"
      }
    }
  },
  methods: {
    vfileAdded(file) {
        var reader = new FileReader();
        reader.readAsArrayBuffer(file);
        reader.onload = () => {
            const data = new Uint8Array(reader.result);
            const xls = XLSX.read(data, {
                type: "array"
            })
            this.xlsxFile = xls
            this.sheets = xls.SheetNames
        }
    },
    vremoved(file) {
        this.sheets = []
    },
    errorInterval() {
        this.showError = true
        setTimeout(function(){ 
            console.log('timeOut')
            this.showError = false
         }.bind(this), 3000);
    },
    calculateTimeDiff(come, back) {
        if (!come) return "00:00"

        const mCome = moment(come, ['h:m a', 'H:m'])
        const mBack = moment(back || '18:00', ['h:m a', 'H:m'])
        if (mCome.hour() >= 18) return come

        const diff = mBack.diff(mCome)
        return moment.utc(diff).format("HH:mm");
    },
    getTimes(times) {
        try {
            let time = times.match(/.{1,5}/g)
            time = time || [null, null]
            time = time.length < 2 ? time.push(null) : time
            
            const come = time[0]
            const back = time[time.length - 1]
            const attend = !come && !back ? 0 : 1
            const diff = this.calculateTimeDiff(come, back)
            const description = ""

            return { come, back, attend, diff, description}
        } catch (e) {
            this.errorInterval()
        }
    },
    convertAttendance(sheet) {
        this.showError = false
        const absen = this.xlsxFile.Sheets[sheet]
        const date = absen['C3'].v
        const dateHeaders = [date, '']
        const headers = ['Nama', 'Total']
        const employeeAttendanceArray = [dateHeaders, headers]
        let colStart = START_NAME_COLUMN

        DAY_ROWS.forEach((row, i) => {
            const day = absen[`${row}4`] !== undefined ? absen[`${row}4`].v : ''
            dateHeaders.push(day, null, null, null, null)
            headers.push('Masuk', 'Pulang', 'Time', '', 'Keterangan')
        })

        while(absen[`K${colStart}`] !== undefined) {
            const name = absen[`K${colStart}`]['v']
            const attArray = []
            let dayCounter = 0
            for (let day of DAY_ROWS) {
                const personAttendance = colStart + 1
                const hoursAttendance = absen[`${day}${personAttendance}`] !== undefined ? absen[`${day}${personAttendance}`].v : ''
                const attendances = this.getTimes(hoursAttendance)
                dayCounter += attendances.attend
                attArray.push(attendances.come, attendances.back, attendances.diff, attendances.attend, attendances.description)
            }

            employeeAttendanceArray.push([name, dayCounter, ...attArray])
            colStart += 2
        }

        let wb = XLSX.utils.book_new();
        wb.SheetNames.push(sheet);
        wb.Sheets[sheet] = XLSX.utils.aoa_to_sheet(employeeAttendanceArray);

        XLSX.writeFile(wb, `${sheet}-${date}_${moment().format('YYYYMMDDHHmmss')}.xlsx`)
    }
  }
};
</script>
<style>
</style>
