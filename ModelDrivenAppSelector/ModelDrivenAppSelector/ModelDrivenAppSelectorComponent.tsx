import * as React from 'react';
import { Dropdown, Option } from '@fluentui/react-components';

export interface IModelDrivenAppSelectorProps {
    selectedAppId: string;
    selectedAppName: string;
    disabled: boolean;
    onChange: (appId: string, appName: string) => void;
}

interface IApp {
    appId: string;
    appName: string;
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

export const ModelDrivenAppSelectorComponent: React.FC<IModelDrivenAppSelectorProps> = (props) => {
    const [apps, setApps] = React.useState<IApp[]>([]);
    const [selectedOption, setSelectedOption] = React.useState<string>(props.selectedAppId ?? '');
    const [selectedDisplayText, setSelectedDisplayText] = React.useState<string>('');
    const [loading, setLoading] = React.useState<boolean>(false);

    const getApps = React.useCallback(
        async (): Promise<IApp[]> => {
            try {
                // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment, @typescript-eslint/no-unsafe-member-access, @typescript-eslint/no-explicit-any, @typescript-eslint/no-unsafe-call
                const baseUrl = (window as any).Xrm?.Page?.context?.getClientUrl?.() ?? '';
                if (!baseUrl) {
                    console.warn('Could not retrieve base URL from Xrm.Page');
                    return [];
                }

                const response = await fetch(
                    `${baseUrl}/api/data/v9.1/appmodules?$select=appmoduleid,name`,
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

                const result = await response.json() as { value: { appmoduleid: string; name: string }[] };
                const appOptions: IApp[] = [];

                if (result.value && Array.isArray(result.value)) {
                    result.value.forEach((app) => {
                        if (app.name) {
                            appOptions.push({
                                appId: app.appmoduleid,
                                appName: app.name
                            });
                        }
                    });
                }

                // Sort alphabetically by name
                appOptions.sort((a, b) => a.appName.localeCompare(b.appName));

                return appOptions;
            } catch (error) {
                console.error('Error retrieving model driven apps:', error);
                return [];
            }
        },
        []
    );

    React.useEffect(() => {
        setLoading(true);
        void getApps().then((result) => {
            setApps(result);
            if (props.selectedAppId) {
                const app = result.find(a => a.appId === props.selectedAppId);
                if (app) {
                    setSelectedDisplayText(`${app.appName} (${app.appId})`);
                }
            }
            setLoading(false);
            return;
        });
    }, [getApps, props.selectedAppId]);

    React.useEffect(() => {
        if (props.selectedAppId && apps.length > 0) {
            const app = apps.find(a => a.appId === props.selectedAppId);
            if (app) {
                setSelectedDisplayText(`${app.appName} (${app.appId})`);
            }
        }
    }, [props.selectedAppId, apps]);

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const handleSelectionChange = (event: any, data: { optionValue: string | undefined }) => {
        const selected = data.optionValue ?? '';
        setSelectedOption(selected);
        const app = apps.find(a => a.appId === selected);
        if (app) {
            setSelectedDisplayText(`${app.appName} (${app.appId})`);
        }
        props.onChange(selected, app?.appName ?? '');
    };

    return (
        <div style={dropdownContainerStyles}>
            <Dropdown
                style={dropdownStyles}
                placeholder={loading ? "Loading model driven apps..." : "Select model driven app"}
                value={selectedDisplayText}
                selectedOptions={[selectedOption]}
                onOptionSelect={handleSelectionChange}
                disabled={props.disabled || loading}
                listbox={{ style: listboxStyles }}
            >
                <Option key="" value="" text="" style={optionStyles} />
                {apps.map((app) => (
                    <Option key={app.appId} value={app.appId} text={`${app.appName} (${app.appId})`} style={optionStyles}>
                        {app.appName} ({app.appId})
                    </Option>
                ))}
            </Dropdown>
        </div>
    );
};
