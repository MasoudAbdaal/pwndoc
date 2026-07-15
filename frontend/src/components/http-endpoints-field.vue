<template>
    <q-field
    ref="field"
    :model-value="rows"
    :label="label"
    stack-label
    outlined
    :hint="hint"
    hide-bottom-space
    :rules="[validateRows]"
    lazy-rules="ondemand"
    >
        <template v-slot:control>
            <div class="full-width q-pt-sm">
                <div v-for="(row, index) in rows" :key="index" class="row q-col-gutter-sm q-mb-sm items-center">
                    <div class="col-12 col-md-2">
                        <q-select
                        :model-value="row.method"
                        :options="methods"
                        :readonly="readonly"
                        label="Method"
                        dense
                        outlined
                        @update:model-value="updateRow(index, 'method', $event)"
                        />
                    </div>
                    <div class="col-12 col-md-5">
                        <q-input
                        :model-value="row.url"
                        :readonly="readonly"
                        label="URL"
                        dense
                        outlined
                        @update:model-value="updateRow(index, 'url', $event)"
                        />
                    </div>
                    <div class="col">
                        <q-input
                        :model-value="row.parameter"
                        :readonly="readonly"
                        label="Parameter"
                        dense
                        outlined
                        @update:model-value="updateRow(index, 'parameter', $event)"
                        />
                    </div>
                    <div v-if="!readonly" class="col-auto">
                        <q-btn flat round dense color="red" icon="delete" aria-label="Delete row" @click="removeRow(index)" />
                    </div>
                </div>
                <q-btn v-if="!readonly" flat dense color="secondary" icon="add" label="Add row" no-caps @click="addRow" />
            </div>
        </template>
        <template v-slot:label>
            {{label}} <span v-if="required" class="text-red">*</span>
        </template>
    </q-field>
</template>

<script>
export default {
    name: 'http-endpoints-field',
    props: {
        modelValue: {
            type: Array,
            default: () => []
        },
        label: {
            type: String,
            default: ''
        },
        hint: {
            type: String,
            default: ''
        },
        required: {
            type: Boolean,
            default: false
        },
        readonly: {
            type: Boolean,
            default: false
        }
    },
    emits: ['update:modelValue'],
    data: function() {
        return {
            methods: ['GET', 'POST', 'PUT', 'PATCH', 'DELETE', 'OPTIONS', 'HEAD']
        }
    },
    computed: {
        rows: function() {
            return Array.isArray(this.modelValue) ? this.modelValue : []
        }
    },
    methods: {
        addRow: function() {
            this.$emit('update:modelValue', [...this.rows, {method: 'GET', url: '', parameter: ''}])
        },
        removeRow: function(index) {
            this.$emit('update:modelValue', this.rows.filter((row, rowIndex) => rowIndex !== index))
        },
        updateRow: function(index, key, value) {
            var rows = this.rows.map((row, rowIndex) => rowIndex === index ? {...row, [key]: value} : row)
            this.$emit('update:modelValue', rows)
        },
        validateRows: function(rows) {
            if (this.required && rows.length === 0)
                return 'Field is required'
            var valid = rows.every(row =>
                row &&
                (!row.method || this.methods.includes(row.method)) &&
                typeof row.url === 'string' &&
                typeof row.parameter === 'string'
            )
            return valid || 'Invalid HTTP affected point'
        },
        validate: function() {
            return this.$refs.field.validate()
        }
    }
}
</script>
