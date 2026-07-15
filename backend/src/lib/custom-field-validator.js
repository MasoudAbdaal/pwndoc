const HTTP_METHODS = ['GET', 'POST', 'PUT', 'PATCH', 'DELETE', 'OPTIONS', 'HEAD']

function validateHttpEndpoints(value) {
    return Array.isArray(value) && value.every(row =>
        row &&
        typeof row === 'object' &&
        (!row.method || HTTP_METHODS.includes(row.method)) &&
        typeof row.url === 'string' &&
        typeof row.parameter === 'string'
    )
}

function validate(customFields) {
    if (!Array.isArray(customFields))
        return false

    return customFields.every(field => {
        var definition = field && field.customField
        var fieldType = definition && typeof definition === 'object' ? definition.fieldType : field && field.fieldType
        return fieldType !== 'http-endpoints' || validateHttpEndpoints(field.text)
    })
}

module.exports = {validate, validateHttpEndpoints, HTTP_METHODS}
