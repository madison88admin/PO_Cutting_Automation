import { getFactoryMappings } from "./db/factoryMapping";
import { getMloMappings } from "./db/mloMapping";
import { getColumnMappings } from "./db/columnMapping";

export async function getFactoryMapping() {
    return await getFactoryMappings();
}

export async function getMloMapping() {
    return await getMloMappings();
}

export async function getColumnMapping(customer?: string) {
    const mappings = await getColumnMappings(customer);

    const result: Record<string, string> = {};
    mappings.forEach(m => {
        // Map: Buy File Column -> Internal Field
        result[m.buy_file_column] = m.internal_field;
    });

    return result;
}

export async function getAllColumnMappings() {
    return await getColumnMappings();
}
