import { Extension } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'

const RTL_RE = /[\u0590-\u08FF\uFB1D-\uFDFF\uFE70-\uFEFF]/
const LTR_RE = /[A-Za-z0-9]/

function getTextDirection(text) {
  for (const char of text || '') {
    if (RTL_RE.test(char)) return 'rtl'
    if (LTR_RE.test(char)) return 'ltr'
  }
  return null
}

function getNodeTextDirection(node) {
  return getTextDirection(node.textContent)
}

export default Extension.create({
  name: 'textDirection',

  addGlobalAttributes() {
    return [
      {
        types: ['paragraph', 'heading', 'listItem'],
        attributes: {
          dir: {
            default: null,
            parseHTML: element => element.getAttribute('dir'),
            renderHTML: attributes => {
              if (!attributes.dir) return {}
              return { dir: attributes.dir }
            },
          },
        },
      },
    ]
  },

  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: new PluginKey('textDirection'),
        appendTransaction: (transactions, oldState, newState) => {
          if (!transactions.some(transaction => transaction.docChanged)) {
            return null
          }

          const tr = newState.tr
          let modified = false

          newState.doc.descendants((node, pos) => {
            if (!['paragraph', 'heading', 'listItem'].includes(node.type.name)) {
              return true
            }

            const dir = getNodeTextDirection(node)
            if (node.attrs.dir !== dir) {
              tr.setNodeMarkup(pos, undefined, { ...node.attrs, dir })
              modified = true
            }

            return true
          })

          return modified ? tr : null
        },
      }),
    ]
  },
})