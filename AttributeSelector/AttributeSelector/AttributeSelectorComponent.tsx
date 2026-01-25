import * as React from 'react';
import { Dropdown, Option } from '@fluentui/react-components';

export interface IAttributeSelectorProps {
    selectedAttribute: string;
    entity: string;
    disabled: boolean;
    onChange: (selectedAttribute: string) => void;
}

interface IAttribute {
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

export const AttributeSelectorComponent: React.FC<IAttributeSelectorProps> = (props) => {
    const [attributes, setAttributes] = React.useState<IAttribute[]>([]);
    const [selectedOption, setSelectedOption] = React.useState<string>(props.selectedAttribute || '');
    const [selectedDisplayText, setSelectedDisplayText] = React.useState<string>('');
    const [loading, setLoading] = React.useState<boolean>(false);

    const getAttributes = React.useCallback(
        async (entityName: string): Promise<IAttribute[]> => {
            try {
                if (!entityName) return [];

                const baseUrl = (window as any).Xrm?.Page?.context?.getClientUrl?.() || '';
                if (!baseUrl) {
                    console.warn('Could not retrieve base URL from Xrm.Page');
                    return [];
                }

                const response = await fetch(
                    `${baseUrl}/api/data/v9.1/EntityDefinitions(LogicalName='${entityName}')/Attributes?$select=LogicalName,DisplayName`,
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
                const attributeOptions: IAttribute[] = [];

                if (result.value && Array.isArray(result.value)) {
                    result.value.forEach((attr: any) => {
                        if (attr.DisplayName?.UserLocalizedLabel?.Label) {
                            attributeOptions.push({
                                key: attr.LogicalName,
                                text: attr.DisplayName.UserLocalizedLabel.Label
                            });
                        }
                    });
                }

                // Sort alphabetically by text (value)
                attributeOptions.sort((a, b) => a.text.localeCompare(b.text));

                return attributeOptions;
            } catch (error) {
                console.error('Error retrieving attributes:', error);
                return [];
            }
        },
        []
    );

    React.useEffect(() => {
        setLoading(true);
        getAttributes(props.entity).then((result) => {
            setAttributes(result);
            setLoading(false);
        });
    }, [props.entity, getAttributes]);

    React.useEffect(() => {
        if (props.selectedAttribute && attributes.length > 0) {
            const attribute = attributes.find(a => a.key === props.selectedAttribute);
            if (attribute) {
                setSelectedDisplayText(`${attribute.text} (${attribute.key})`);
            }
        }
    }, [props.selectedAttribute, attributes]);

    const handleSelectionChange = (event: any, data: any) => {
        const selected = data.optionValue || '';
        setSelectedOption(selected);
        const attribute = attributes.find(a => a.key === selected);
        if (attribute) {
            setSelectedDisplayText(`${attribute.text} (${attribute.key})`);
        }
        props.onChange(selected);
    };

    return (
        <div style={dropdownContainerStyles}>
            <Dropdown
                style={dropdownStyles}
                placeholder={loading ? "Loading attributes..." : "Select attribute"}
                value={selectedDisplayText}
                selectedOptions={[selectedOption]}
                onOptionSelect={handleSelectionChange}
                disabled={props.disabled || loading}
                listbox={{ style: listboxStyles }}
            >
                <Option key="" value="" text="" style={optionStyles} />
                {attributes.map((attr) => (
                    <Option key={attr.key} value={attr.key} text={`${attr.text} (${attr.key})`} style={optionStyles}>
                        {attr.text} ({attr.key})
                    </Option>
                ))}
            </Dropdown>
        </div>
    );
};
