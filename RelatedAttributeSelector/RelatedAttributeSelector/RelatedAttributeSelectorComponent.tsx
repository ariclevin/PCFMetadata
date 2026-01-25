import * as React from 'react';
import { Dropdown, Option } from '@fluentui/react-components';

export interface IRelatedAttributeSelectorProps {
    selectedAttribute: string;
    formEntity: string;
    queryEntity: string;
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

export const RelatedAttributeSelectorComponent: React.FC<IRelatedAttributeSelectorProps> = (props) => {
    const [lookups, setLookups] = React.useState<ILookup[]>([]);
    const [selectedOption, setSelectedOption] = React.useState<string>(props.selectedAttribute || '');
    const [selectedDisplayText, setSelectedDisplayText] = React.useState<string>('');
    const [loading, setLoading] = React.useState<boolean>(false);

    const getRelatedLookups = React.useCallback(
        async (formEntity: string, queryEntity: string): Promise<ILookup[]> => {
            try {
                if (!formEntity || !queryEntity) return [];

                const baseUrl = (window as any).Xrm?.Page?.context?.getClientUrl?.() || '';
                if (!baseUrl) {
                    console.warn('Could not retrieve base URL from Xrm.Page');
                    return [];
                }

                const response = await fetch(
                    `${baseUrl}/api/data/v9.1/EntityDefinitions(LogicalName='${formEntity}')/Attributes/Microsoft.Dynamics.CRM.LookupAttributeMetadata?$select=LogicalName,DisplayName,Targets`,
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
                            // Filter to only lookups that target the queryEntity
                            if (lookup.Targets && Array.isArray(lookup.Targets) && lookup.Targets.includes(queryEntity)) {
                                lookupOptions.push({
                                    key: lookup.LogicalName,
                                    text: lookup.DisplayName.UserLocalizedLabel.Label
                                });
                            }
                        }
                    });
                }

                // Sort alphabetically by text
                lookupOptions.sort((a, b) => a.text.localeCompare(b.text));

                return lookupOptions;
            } catch (error) {
                console.error('Error retrieving related lookups:', error);
                return [];
            }
        },
        []
    );

    React.useEffect(() => {
        setLoading(true);
        getRelatedLookups(props.formEntity, props.queryEntity).then((result) => {
            setLookups(result);
            setLoading(false);
        });
    }, [props.formEntity, props.queryEntity, getRelatedLookups]);

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
                placeholder={loading ? "Loading related lookups..." : "Select related lookup"}
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
