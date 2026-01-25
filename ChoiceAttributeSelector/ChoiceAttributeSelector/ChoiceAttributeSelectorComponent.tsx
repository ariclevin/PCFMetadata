import * as React from 'react';
import { Dropdown, Option } from '@fluentui/react-components';

export interface IChoiceAttributeSelectorProps {
    selectedAttribute: string;
    entity: string;
    disabled: boolean;
    onChange: (selectedAttribute: string) => void;
}

interface IChoice {
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

export const ChoiceAttributeSelectorComponent: React.FC<IChoiceAttributeSelectorProps> = (props) => {
    const [choiceAttributes, setChoiceAttributes] = React.useState<IChoice[]>([]);
    const [selectedOption, setSelectedOption] = React.useState<string>(props.selectedAttribute || '');
    const [selectedDisplayText, setSelectedDisplayText] = React.useState<string>('');
    const [loading, setLoading] = React.useState<boolean>(false);

    const getChoiceAttributes = React.useCallback(
        async (entityName: string): Promise<IChoice[]> => {
            try {
                if (!entityName) return [];

                const baseUrl = (window as any).Xrm?.Page?.context?.getClientUrl?.() || '';
                if (!baseUrl) {
                    console.warn('Could not retrieve base URL from Xrm.Page');
                    return [];
                }

                const response = await fetch(
                    `${baseUrl}/api/data/v9.1/EntityDefinitions(LogicalName='${entityName}')/Attributes?$select=LogicalName,DisplayName,AttributeType&$filter=AttributeType eq 'Picklist' or AttributeType eq 'Status' or AttributeType eq 'State'`,
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
                const choiceAttributes: IChoice[] = [];

                if (result.value && Array.isArray(result.value)) {
                    result.value.forEach((optionSet: any) => {
                        if (optionSet.DisplayName?.UserLocalizedLabel?.Label) {
                            choiceAttributes.push({
                                key: optionSet.LogicalName,
                                text: optionSet.DisplayName.UserLocalizedLabel.Label
                            });
                        }
                    });
                }

                // Sort alphabetically by text
                choiceAttributes.sort((a, b) => a.text.localeCompare(b.text));
                return choiceAttributes;
            } catch (error) {
                console.error('Error retrieving option sets:', error);
                return [];
            }
        },
        []
    );

    React.useEffect(() => {
        setLoading(true);
        getChoiceAttributes(props.entity).then((result) => {
            setChoiceAttributes(result);
            setLoading(false);
        });
    }, [props.entity, getChoiceAttributes]);

    React.useEffect(() => {
        if (props.selectedAttribute && choiceAttributes.length > 0) {
            const choice = choiceAttributes.find(c => c.key === props.selectedAttribute);
            if (choice) {
                setSelectedDisplayText(`${choice.text} (${choice.key})`);
            }
        }
    }, [props.selectedAttribute, choiceAttributes]);

    const handleSelectionChange = (event: any, data: any) => {
        const selected = data.optionValue || '';
        setSelectedOption(selected);
        const choice = choiceAttributes.find(c => c.key === selected);
        if (choice) {
            setSelectedDisplayText(`${choice.text} (${choice.key})`);
        }
        props.onChange(selected);
    };

    return (
        <div style={dropdownContainerStyles}>
            <Dropdown
                style={dropdownStyles}
                placeholder={loading ? "Loading choices..." : "Select choice"}
                value={selectedDisplayText}
                selectedOptions={[selectedOption]}
                onOptionSelect={handleSelectionChange}
                disabled={props.disabled || loading}
                listbox={{ style: listboxStyles }}
            >
                <Option key="" value="" text="" style={optionStyles} />
                {choiceAttributes.map((choiceAttribute) => (
                    <Option key={choiceAttribute.key} value={choiceAttribute.key} text={`${choiceAttribute.text} (${choiceAttribute.key})`} style={optionStyles}>
                        {choiceAttribute.text} ({choiceAttribute.key})
                    </Option>
                ))}
            </Dropdown>
        </div>
    );
};
