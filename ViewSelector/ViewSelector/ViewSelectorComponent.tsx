import * as React from 'react';
import { Dropdown, Option } from '@fluentui/react-components';

export interface IViewSelectorProps {
    selectedView: string;
    entity: string;
    disabled: boolean;
    onChange: (selectedView: string) => void;
}

interface IView {
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

export const ViewSelectorComponent: React.FC<IViewSelectorProps> = (props) => {
    const [views, setViews] = React.useState<IView[]>([]);
    const [selectedOption, setSelectedOption] = React.useState<string>(props.selectedView || '');
    const [selectedDisplayText, setSelectedDisplayText] = React.useState<string>('');
    const [loading, setLoading] = React.useState<boolean>(false);

    const getViews = React.useCallback(
        async (entityName: string): Promise<IView[]> => {
            try {
                if (!entityName) return [];

                const baseUrl = (window as any).Xrm?.Page?.context?.getClientUrl?.() || '';
                if (!baseUrl) {
                    console.warn('Could not retrieve base URL from Xrm.Page');
                    return [];
                }

                const response = await fetch(
                    `${baseUrl}/api/data/v9.1/savedqueries?$select=savedqueryid,name,returnedtypecode&$filter=returnedtypecode eq '${entityName}'`,
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
                const viewOptions: IView[] = [];

                if (result.value && Array.isArray(result.value)) {
                    result.value.forEach((view: any) => {
                        if (view.name) {
                            viewOptions.push({
                                key: view.savedqueryid,
                                text: view.name
                            });
                        }
                    });
                }

                // Sort alphabetically by text
                viewOptions.sort((a, b) => a.text.localeCompare(b.text));

                return viewOptions;
            } catch (error) {
                console.error('Error retrieving views:', error);
                return [];
            }
        },
        []
    );

    React.useEffect(() => {
        setLoading(true);
        getViews(props.entity).then((result) => {
            setViews(result);
            setLoading(false);
        });
    }, [props.entity, getViews]);

    React.useEffect(() => {
        if (props.selectedView && views.length > 0) {
            const view = views.find(v => v.key === props.selectedView);
            if (view) {
                setSelectedDisplayText(`${view.text} (${view.key})`);
            }
        }
    }, [props.selectedView, views]);

    const handleSelectionChange = (event: any, data: any) => {
        const selected = data.optionValue || '';
        setSelectedOption(selected);
        const view = views.find(v => v.key === selected);
        if (view) {
            setSelectedDisplayText(`${view.text} (${view.key})`);
        }
        props.onChange(selected);
    };

    return (
        <div style={dropdownContainerStyles}>
            <Dropdown
                style={dropdownStyles}
                placeholder={loading ? "Loading views..." : "Select view"}
                value={selectedDisplayText}
                selectedOptions={[selectedOption]}
                onOptionSelect={handleSelectionChange}
                disabled={props.disabled || loading}
                listbox={{ style: listboxStyles }}
            >
                <Option key="" value="" text="" style={optionStyles} />
                {views.map((view) => (
                    <Option key={view.key} value={view.key} text={`${view.text} (${view.key})`} style={optionStyles}>
                        {view.text} ({view.key})
                    </Option>
                ))}
            </Dropdown>
        </div>
    );
};
