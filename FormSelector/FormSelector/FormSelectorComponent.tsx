import * as React from 'react';
import { Dropdown, Option } from '@fluentui/react-components';

export interface IFormSelectorProps {
    selectedForm: string;
    entity: string;
    disabled: boolean;
    onChange: (selectedForm: string) => void;
}

interface IForm {
    key: string;
    text: string;
}

const dropdownStyles: React.CSSProperties = {
    width: '100%',
    minWidth: '200px'
};

const dropdownContainerStyles: React.CSSProperties = {
    display: 'flex',
    flexDirection: 'column',
    width: '100%'
};

const listboxStyles: React.CSSProperties = {
    maxHeight: '300px',
    overflow: 'auto',
    backgroundColor: '#ffffff',
    border: '1px solid #e0e0e0',
    boxShadow: '0 2px 8px rgba(0, 0, 0, 0.1)',
    padding: '4px 0'
};

const optionStyles: React.CSSProperties = {
    padding: '10px 12px',
    lineHeight: '1.6'
};

export const FormSelectorComponent: React.FC<IFormSelectorProps> = (props) => {
    const [forms, setForms] = React.useState<IForm[]>([]);
    const [selectedOption, setSelectedOption] = React.useState<string>(props.selectedForm || '');
    const [selectedDisplayText, setSelectedDisplayText] = React.useState<string>('');
    const [loading, setLoading] = React.useState<boolean>(false);

    const getForms = React.useCallback(
        async (entityName: string): Promise<IForm[]> => {
            try {
                if (!entityName) return [];

                const baseUrl = (window as any).Xrm?.Page?.context?.getClientUrl?.() || '';
                if (!baseUrl) {
                    console.warn('Could not retrieve base URL from Xrm.Page');
                    return [];
                }

                const response = await fetch(
                    `${baseUrl}/api/data/v9.1/systemforms?$filter=objecttypecode eq '${entityName}'&$select=formid,name`,
                    {
                        method: 'GET',
                        headers: {
                            'OData-MaxVersion': '4.0',
                            'OData-Version': '4.0',
                            'Accept': 'application/json',
                            'Content-Type': 'application/json; charset=utf-8'
                        }
                    }
                );

                if (!response.ok) {
                    throw new Error(`HTTP error! status: ${response.status}`);
                }

                const result = await response.json();
                const formOptions: IForm[] = [];

                if (result.value && Array.isArray(result.value)) {
                    result.value.forEach((form: any) => {
                        formOptions.push({
                            key: form.formid,
                            text: form.name
                        });
                    });
                }

                // Sort alphabetically by text
                formOptions.sort((a, b) => a.text.localeCompare(b.text));

                return formOptions;
            } catch (error) {
                console.error('Error retrieving forms:', error);
                return [];
            }
        },
        []
    );

    React.useEffect(() => {
        setLoading(true);
        getForms(props.entity).then((result) => {
            setForms(result);
            setLoading(false);
        });
    }, [props.entity, getForms]);

    React.useEffect(() => {
        if (props.selectedForm && forms.length > 0) {
            const form = forms.find(f => f.key === props.selectedForm);
            if (form) {
                setSelectedDisplayText(`${form.text} (${form.key})`);
            }
        }
    }, [props.selectedForm, forms]);

    const handleSelectionChange = (event: any, data: any) => {
        const selected = data.optionValue || '';
        setSelectedOption(selected);
        const form = forms.find(f => f.key === selected);
        if (form) {
            setSelectedDisplayText(`${form.text} (${form.key})`);
        }
        props.onChange(selected);
    };

    return (
        <div style={dropdownContainerStyles}>
            <Dropdown
                style={dropdownStyles}
                placeholder={loading ? "Loading forms..." : "Select form"}
                value={selectedDisplayText}
                selectedOptions={[selectedOption]}
                onOptionSelect={handleSelectionChange}
                disabled={props.disabled || loading}
                listbox={{ style: listboxStyles }}
            >
                <Option key="" value="" text="" style={optionStyles} />
                {forms.map((form) => (
                    <Option key={form.key} value={form.key} text={`${form.text} (${form.key})`} style={optionStyles}>
                        {form.text} ({form.key})
                    </Option>
                ))}
            </Dropdown>
        </div>
    );
};
