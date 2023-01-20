export const escapeSingleQuotesForFilter = (str: string) => {
    return str.replace(/'/g, "''");
};

export const escapeSingleQuotesForSearch = (str: string) => {
    return str.replace(/'/g, "\'");
};