import * as React from 'react';
import { Dropdown, Option } from '@fluentui/react-components';

export interface ILookupAttributeSelectorProps {
    selectedAttribute: string;
    entity: string;
    disabled: boolean;
    onChange: (selectedAttribute: string) => void;
}

interface ILookup {
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

export const LookupAttributeSelectorComponent: React.FC<ILookupAttributeSelectorProps> = (props) => {
    const [lookups, setLookups] = React.useState<ILookup[]>([]);
    const [selectedOption, setSelectedOption] = React.useState<string>(props.selectedAttribute || '');
    const [selectedDisplayText, setSelectedDisplayText] = React.useState<string>('');
    const [loading, setLoading] = React.useState<boolean>(false);

    const getLookups = React.useCallback(
        async (entityName: string): Promise<ILookup[]> => {
            try {
                if (!entityName) return [];

                const baseUrl = (window as any).Xrm?.Page?.context?.getClientUrl?.() || '';
                if (!baseUrl) {
                    console.warn('Could not retrieve base URL from Xrm.Page');
                    return [];
                }

                const response = await fetch(
                    `${baseUrl}/api/data/v9.1/EntityDefinitions(LogicalName='${entityName}')/Attributes?$select=LogicalName,DisplayName,AttributeType&$filter=AttributeType eq 'Lookup'`,
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
                const lookupOptions: ILookup[] = [];

                if (result.value && Array.isArray(result.value)) {
                    result.value.forEach((lookup: any) => {
                        if (lookup.DisplayName?.UserLocalizedLabel?.Label) {
                            lookupOptions.push({
                                key: lookup.LogicalName,
                                text: lookup.DisplayName.UserLocalizedLabel.Label
                            });
                        }
                    });
                }

                // Sort alphabetically by text
                lookupOptions.sort((a, b) => a.text.localeCompare(b.text));

                return lookupOptions;
            } catch (error) {
                console.error('Error retrieving lookups:', error);
                return [];
            }
        },
        []
    );

    React.useEffect(() => {
        setLoading(true);
        getLookups(props.entity).then((result) => {
            setLookups(result);
            setLoading(false);
        });
    }, [props.entity, getLookups]);

    React.useEffect(() => {
        if (props.selectedAttribute && lookups.length > 0) {
            const lookup = lookups.find(l => l.key === props.selectedAttribute);
            if (lookup) {
                setSelectedDisplayText(`${lookup.text} (${lookup.key})`);
            }
        }
    }, [props.selectedAttribute, lookups]);

    const handleSelectionChange = (event: any, data: any) => {
        const selected = data.optionValue || '';
        setSelectedOption(selected);
        const lookup = lookups.find(l => l.key === selected);
        if (lookup) {
            setSelectedDisplayText(`${lookup.text} (${lookup.key})`);
        }
        props.onChange(selected);
    };

    return (
        <div style={dropdownContainerStyles}>
            <Dropdown
                style={dropdownStyles}
                placeholder={loading ? "Loading lookups..." : "Select lookup"}
                value={selectedDisplayText}
                selectedOptions={[selectedOption]}
                onOptionSelect={handleSelectionChange}
                disabled={props.disabled || loading}
                listbox={{ style: listboxStyles }}
            >
                <Option key="" value="" text="" style={optionStyles} />
                {lookups.map((lookup) => (
                    <Option key={lookup.key} value={lookup.key} text={`${lookup.text} (${lookup.key})`} style={optionStyles}>
                        {lookup.text} ({lookup.key})
                    </Option>
                ))}
            </Dropdown>
        </div>
    );
};
