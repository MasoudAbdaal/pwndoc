import { describe, it, expect } from 'vitest'
import { createTestWrapper } from '../../test-utils'
import HttpEndpointsField from '@/components/http-endpoints-field.vue'

function createWrapper(modelValue = []) {
  return createTestWrapper(HttpEndpointsField, {
    props: { modelValue, label: 'Affected Endpoints' },
    global: {
      stubs: {
        'q-field': { template: '<div><slot name="control" /></div>', methods: { validate: () => true } },
        'q-select': true,
        'q-input': true,
        'q-btn': true
      }
    }
  })
}

describe('HttpEndpointsField', () => {
  it('adds, edits, and removes rows through model updates', async () => {
    const wrapper = createWrapper()

    wrapper.vm.addRow()
    const addedRows = wrapper.emitted('update:modelValue')[0][0]
    expect(addedRows).toEqual([
      { method: 'GET', url: '', parameter: '' }
    ])
    await wrapper.setProps({ modelValue: addedRows })

    wrapper.vm.updateRow(0, 'url', '/api/users')
    const editedRows = wrapper.emitted('update:modelValue')[1][0]
    expect(editedRows).toEqual([
      { method: 'GET', url: '/api/users', parameter: '' }
    ])
    await wrapper.setProps({ modelValue: editedRows })

    wrapper.vm.removeRow(0)
    expect(wrapper.emitted('update:modelValue')[2][0]).toEqual([])
  })

  it('validates methods and string values without strict URL validation', () => {
    const wrapper = createWrapper()
    expect(wrapper.vm.validateRows([
      { method: 'POST', url: 'https://example.com/api/login', parameter: 'username' }
    ])).toBe(true)
    expect(wrapper.vm.validateRows([
      { method: 'TRACE', url: '/api/users', parameter: 'id' }
    ])).toBe('Invalid HTTP affected point')
    expect(wrapper.vm.validateRows([
      { method: 'GET', url: 42, parameter: 'id' }
    ])).toBe('Invalid HTTP affected point')
  })
})
