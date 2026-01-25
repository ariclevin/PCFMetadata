import * as React from 'react';
import { Dropdown, Option } from '@fluentui/react-components';
import type { DropdownProps } from "@fluentui/react-components";

export interface IEntitySelectorProps {
    selectedLogicalName: string;
    selectedDisplayName: string;
    activitiesOnly: string;
    supportsActivities: string;
    disabled: boolean;
    onChange: (logicalName: string, displayName: string) => void;
}

interface IEntity {
    logicalName: string;
    displayName: string;
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

export const EntitySelectorComponent: React.FC<IEntitySelectorProps> = (props) => {
    const [entities, setEntities] = React.useState<IEntity[]>([]);
    const [selectedOption, setSelectedOption] = React.useState<string>(props.selectedLogicalName || '');
    const [selectedDisplayText, setSelectedDisplayText] = React.useState<string>('');
    const [loading, setLoading] = React.useState<boolean>(false);

    const getEntities = React.useCallback(
        async (onlyActivities: string, supportsActivities: string): Promise<IEntity[]> => {
            try {
                const baseUrl = (window as any).Xrm?.Page?.context?.getClientUrl?.() || '';
                if (!baseUrl) {
                    console.warn('Could not retrieve base URL from Xrm.Page');
                    return [];
                }

                let filter = "IsValidForAdvancedFind%20eq%20true%20and%20IsCustomizable/Value%20eq%20true";
                if (onlyActivities !== "All") {
                    filter += "%20and%20IsActivity%20eq%20true";
                }
                if (supportsActivities !== "All") {
                    filter += "%20and%20HasActivities%20eq%20true";
                }

                const response = await fetch(
                    `${baseUrl}/api/data/v9.1/EntityDefinitions?$select=LogicalName,DisplayName&$filter=${filter}`,
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
                const entityOptions: IEntity[] = [];

                if (result.value && Array.isArray(result.value)) {
                    result.value.forEach((entity: any) => {
                        if (entity.DisplayName?.UserLocalizedLabel?.Label) {
                            entityOptions.push({
                                logicalName: entity.LogicalName,
                                displayName: entity.DisplayName.UserLocalizedLabel.Label
                            });
                        }
                    });
                }

                // Sort alphabetically by display name
                entityOptions.sort((a, b) => a.displayName.localeCompare(b.displayName));

                return entityOptions;
            } catch (error) {
                console.error('Error retrieving entities:', error);
                return [];
            }
        },
        []
    );

    React.useEffect(() => {
        setLoading(true);
        getEntities(props.activitiesOnly, props.supportsActivities).then((result) => {
            setEntities(result);
            if (props.selectedLogicalName) {
                const entity = result.find(e => e.logicalName === props.selectedLogicalName);
                if (entity) {
                    setSelectedDisplayText(`${entity.displayName} (${entity.logicalName})`);
                }
            }
            setLoading(false);
        });
    }, [props.activitiesOnly, props.supportsActivities, getEntities]);

    React.useEffect(() => {
        if (props.selectedLogicalName && entities.length > 0) {
            const entity = entities.find(e => e.logicalName === props.selectedLogicalName);
            if (entity) {
                setSelectedDisplayText(`${entity.displayName} (${entity.logicalName})`);
            }
        }
    }, [props.selectedLogicalName, entities]);

    const handleSelectionChange = (event: any, data: any) => {
        const selected = data.optionValue || '';
        setSelectedOption(selected);
        const entity = entities.find(e => e.logicalName === selected);
        if (entity) {
            setSelectedDisplayText(`${entity.displayName} (${entity.logicalName})`);
        }
        props.onChange(selected, entity?.displayName || '');
    };

    return (
        <div style={dropdownContainerStyles}>
            <Dropdown
                style={dropdownStyles}
                placeholder={loading ? "Loading entities..." : "Select entity"}
                value={selectedDisplayText}
                selectedOptions={[selectedOption]}
                onOptionSelect={handleSelectionChange}
                disabled={props.disabled || loading}
                listbox={{ style: listboxStyles }}
            >

                <Option key="" value="" text="" style={optionStyles} />
                {entities.map((entity) => (
                    <Option key={entity.logicalName} value={entity.logicalName} text={`${entity.displayName} (${entity.logicalName})`} style={optionStyles}>
                        {entity.displayName} ({entity.logicalName})
                    </Option>
                ))}
            </Dropdown>
        </div>
    );
};
