import { useEffect, useState, useMemo, useCallback } from 'react';
import { apipms } from '../../../../../service/apipms';
import { useToast } from '../../../../../context/ToastContext';

export const usePartNSupplierEmp = () => {
    const { showToast } = useToast();

    const [suppliers, setSuppliers] = useState([]);
    const [partNumbers, setPartNumbers] = useState([]);
    const [itemCodeSuppliers, setItemCodeSuppliers] = useState([]);

    const [selectedSupplier, setSelectedSupplier] = useState('');
    const [selectedPartNumber, setSelectedPartNumber] = useState('');

    const [itemCodeSupplier, setItemCodeSupplier] = useState('');
    const [descriptionES, setDescriptionES] = useState('');
    const [descriptionEN, setDescriptionEN] = useState('');
    const [fabricContent, setFabricContent] = useState('');
    const [hts, setHts] = useState('');
    const [width, setWidth] = useState('');
    const [weight, setWeight] = useState('');

    const [items, setItems] = useState([]);
    const [errors, setErrors] = useState({});
    const [saving, setSaving] = useState(false);
    const [editItemCode, setEditItemCode] = useState(null);
    const [loadingEdit, setLoadingEdit] = useState(false);
    const [selectedItemIndex, setSelectedItemIndex] = useState(null);
    const [searchItemCode, setSearchItemCode] = useState('');


    const partNumberLabel = (pn) => `${pn.codePAH}`;

    const fetchAllItemCodeSuppliers = useCallback(async () => {
        try {
            const { data } = await apipms.get(
                '/filemaintenance/viewsupplieritemcode'
            );

            setSuppliers(data.suppliers || []);
            setPartNumbers(data.partNumbers || []);
            setItemCodeSuppliers(data.itemCodeSuppliers || []);
        } catch (error) {
            console.error('Error al cargar ItemCodeSupplier', error);
            const message =
                error.response?.data?.message || 'Error al cargar datos';
            showToast('error', message);
        }
    }, [showToast]);

    useEffect(() => {
        fetchAllItemCodeSuppliers();
    }, [fetchAllItemCodeSuppliers]);

    const validateRow = () => {
        const newErrors = {};

        if (!selectedSupplier) newErrors.supplier = 'Requerido';
        if (!selectedPartNumber) newErrors.partNumber = 'Requerido';
        if (!itemCodeSupplier.trim()) newErrors.itemCodeSupplier = 'Requerido';
        if (!descriptionEN.trim()) newErrors.descriptionEN = 'Requerido';

        setErrors(newErrors);
        return Object.keys(newErrors).length === 0;
    };

    const handleAddToList = () => {
        if (!validateRow()) return;

        const supplierObj = suppliers.find(
            (s) => s.supplierID === selectedSupplier
        );
        const partNumberObj = partNumbers.find(
            (p) => p.partNumberID === selectedPartNumber
        );

        setItems((prev) => [
            ...prev,
            {
                supplierID: selectedSupplier,
                supplierName: supplierObj?.brandName || supplierObj?.legalName,
                partNumberID: selectedPartNumber,
                partNumberLabel: partNumberLabel(partNumberObj),

                itemCodeSupplier,
                descriptionES,
                descriptionEN,
                fabricContent,
                hts,
                width,
                weight,
            },
        ]);

        // limpiar selects 
        setSelectedSupplier('');
        setSelectedPartNumber('');
        setSelectedItemIndex(null);

        // limpiar detalles
        setItemCodeSupplier('');
        setDescriptionES('');
        setDescriptionEN('');
        setFabricContent('');
        setHts('');
        setWidth('');
        setWeight('');
        setErrors({});
    };

    const handleRemoveItem = (indexToRemove) => {

        showToast('info', 'Ítem eliminado de la lista.');

        setItems((prev) => prev.filter((_, index) => index !== indexToRemove));

        // si estaba seleccionada, quitar selección
        setSelectedItemIndex((current) =>
            current === indexToRemove ? null : current
        );
    };

    const handleSelectItemRow = (row, index) => {
        setSelectedItemIndex(index);

        // seleccionar también PartNumber y Proveedor
        setSelectedPartNumber(row.partNumberID);
        setSelectedSupplier(row.supplierID);

        // cargar los detalles en los TextField de la derecha
        setItemCodeSupplier(row.itemCodeSupplier || '');
        setDescriptionES(row.descriptionES || '');
        setDescriptionEN(row.descriptionEN || '');
        setFabricContent(row.fabricContent || '');
        setHts(row.hts || '');
        setWidth(row.width || '');
        setWeight(row.weight || '');
    };

    const partNumberCodeMap = useMemo(() => {
        const map = new Map();
        partNumbers.forEach((p) => {
            map.set(p.partNumberID, (p.codePAH || '').toLowerCase());
        });
        return map;
    }, [partNumbers]);


    const handleSaveAll = async () => {
        if (items.length === 0) return;

        setSaving(true);
        try {
            const responses = await Promise.all(
                items.map((item) =>
                    apipms.post('/filemaintenance/itemcodesupplier', {
                        supplierID: item.supplierID,
                        partNumberID: item.partNumberID,
                        itemCodeSupplier: item.itemCodeSupplier,
                        descriptionES: item.descriptionES,
                        descriptionEN: item.descriptionEN,
                        fabricContent: item.fabricContent,
                        hts: item.hts,
                        width: item.width,
                        weight: item.weight,
                    })
                )
            );

            const firstMsg = responses[0]?.data?.message;
            if (firstMsg) {
                showToast('success', firstMsg);
            }

            // reset general
            setSelectedSupplier('');
            setSelectedPartNumber('');
            setItems([]);
            setItemCodeSupplier('');
            setDescriptionES('');
            setDescriptionEN('');
            setFabricContent('');
            setHts('');
            setWidth('');
            setWeight('');
            setErrors({});
            setSelectedItemIndex(null);

            // recargar tabla grande
            await fetchAllItemCodeSuppliers();
        } catch (error) {
            console.error(
                'Error al guardar registros',
                error.response?.data || error
            );
            const message =
                error.response?.data?.message || 'Error al guardar registros';
            showToast('error', message);
        } finally {
            setSaving(false);
        }
    };

    const canAddRow = !!(selectedSupplier && selectedPartNumber);

    const getSupplierName = useCallback(
        (id) => {
            const s = suppliers.find((sup) => sup.supplierID === id);
            return s ? s.brandName || s.legalName : id;
        },
        [suppliers]
    );

    const getPartNumberCode = useCallback(
        (id) => {
            const p = partNumbers.find((pn) => pn.partNumberID === id);
            return p ? p.codePAH : id;
        },
        [partNumbers]
    );

    const handleStartEdit = useCallback(
        (row) => {
            setEditItemCode(row.itemCodeSupplier);
            setSelectedSupplier(row.supplierID);
            setSelectedPartNumber(row.partNumberID);
            setItemCodeSupplier(row.itemCodeSupplier || '');
            setDescriptionES(row.descriptionES || '');
            setDescriptionEN(row.descriptionEN || '');
            setFabricContent(row.fabricContent || '');
            setHts(row.hts || '');
            setWidth(row.width || '');
            setWeight(row.weight || '');
            setErrors({});
        },
        [
            setEditItemCode,
            setSelectedSupplier,
            setSelectedPartNumber,
            setItemCodeSupplier,
            setDescriptionES,
            setDescriptionEN,
            setFabricContent,
            setHts,
            setWidth,
            setWeight,
            setErrors,
        ]
    );

    const handleCancelEdit = () => {
        setEditItemCode(null);
        setSelectedSupplier('');
        setSelectedPartNumber('');
        setItemCodeSupplier('');
        setDescriptionES('');
        setDescriptionEN('');
        setFabricContent('');
        setHts('');
        setWidth('');
        setWeight('');
        setErrors({});
    };

    const handleSaveEdit = async () => {
        if (!editItemCode) return;
        if (!validateRow()) return;

        setLoadingEdit(true);
        try {
            const payload = {
                supplierID: selectedSupplier,
                partNumberID: selectedPartNumber,
                itemCodeSupplier,
                descriptionES,
                descriptionEN,
                fabricContent,
                hts,
                width,
                weight,
            };

            const { data } = await apipms.put(
                `/filemaintenance/updatesupplieritemcode/${encodeURIComponent(
                    editItemCode
                )}`,
                payload
            );

            const message =
                data?.message ||
                'Item Code Supplier actualizado correctamente';
            const type = data?.ok === false ? 'warning' : 'success';
            showToast(type, message);

            if (data?.ok !== false) {
                setItemCodeSuppliers((prev) =>
                    prev.map((row) =>
                        row.itemCodeSupplier === editItemCode
                            ? {
                                ...row,
                                supplierID: payload.supplierID,
                                partNumberID: payload.partNumberID,
                                itemCodeSupplier: payload.itemCodeSupplier,
                                descriptionES: payload.descriptionES,
                                descriptionEN: payload.descriptionEN,
                                fabricContent: payload.fabricContent,
                                hts: payload.hts,
                                width: payload.width,
                                weight: payload.weight,
                            }
                            : row
                    )
                );

                setEditItemCode(null);
                setSelectedSupplier('');
                setSelectedPartNumber('');
                setItemCodeSupplier('');
                setDescriptionES('');
                setDescriptionEN('');
                setFabricContent('');
                setHts('');
                setWidth('');
                setWeight('');
                setErrors({});
            }
        } catch (error) {
            console.error('Error actualizando ItemCodeSupplier', error);
            const message =
                error.response?.data?.message ||
                'Error actualizando el Item Code Supplier';
            showToast('error', message);
        } finally {
            setLoadingEdit(false);
        }
    };

    const handleDeleteItemCode = useCallback(
        async (itemCode) => {
            try {
                const { data } = await apipms.delete(
                    `/filemaintenance/deleteitemcode/${encodeURIComponent(
                        itemCode
                    )}`
                );

                if (data?.ok) {
                    showToast('success', data.message || 'Item Code eliminado.');
                    await fetchAllItemCodeSuppliers();
                } else {
                    showToast(
                        'warning',
                        data?.message || 'No se pudo eliminar el Item Code.'
                    );
                }
            } catch (error) {
                const backendMessage = error.response?.data?.message;
                showToast(
                    'error',
                    backendMessage || 'Error al eliminar el Item Code.'
                );
            }
        },
        [fetchAllItemCodeSuppliers, showToast]
    );

    const filteredRows = useMemo(() => {
        const normalized = searchItemCode.trim().toLowerCase();
        if (!normalized) return itemCodeSuppliers;

        return itemCodeSuppliers.filter((row) => {
            const itemCode = (row.itemCodeSupplier || '').toLowerCase();
            const partNumberCode = partNumberCodeMap.get(row.partNumberID) || '';

            return (
                itemCode.includes(normalized) ||
                partNumberCode.includes(normalized)
            );
        });
    }, [searchItemCode, itemCodeSuppliers, partNumberCodeMap]);


    return {
        suppliers,
        partNumbers,
        itemCodeSuppliers,

        selectedSupplier,
        setSelectedSupplier,
        selectedPartNumber,
        setSelectedPartNumber,

        itemCodeSupplier,
        setItemCodeSupplier,
        descriptionES,
        setDescriptionES,
        descriptionEN,
        setDescriptionEN,
        fabricContent,
        setFabricContent,
        hts,
        setHts,
        width,
        setWidth,
        weight,
        setWeight,

        items,
        errors,
        saving,

        editItemCode,
        loadingEdit,
        selectedItemIndex,

        searchItemCode,
        setSearchItemCode,
        handleSelectItemRow,

        handleAddToList,
        handleRemoveItem,
        handleSaveAll,
        canAddRow,
        getSupplierName,
        getPartNumberCode,
        handleStartEdit,
        handleCancelEdit,
        handleSaveEdit,
        handleDeleteItemCode,
        filteredRows,
    };
};
