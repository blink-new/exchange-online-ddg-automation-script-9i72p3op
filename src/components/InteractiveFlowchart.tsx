import React, { useState, useRef } from 'react';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from './ui/card';
import { Badge } from './ui/badge';
import { Button } from './ui/button';
import { Separator } from './ui/separator';
import { 
  Play, 
  Square, 
  Diamond, 
  Circle, 
  XCircle, 
  CheckCircle, 
  Settings, 
  ZoomIn, 
  ZoomOut, 
  RotateCcw,
  Code,
  Database,
  Shield,
  FileText,
  Target,
  GitBranch
} from 'lucide-react';

interface FlowNode {
  id: string;
  type: 'start' | 'end' | 'process' | 'decision' | 'function' | 'error' | 'success';
  title: string;
  description: string;
  x: number;
  y: number;
  width: number;
  height: number;
  connections: string[];
  details: {
    code?: string;
    variables?: string[];
    functions?: string[];
    errorHandling?: string;
    purpose: string;
  };
}

const InteractiveFlowchart: React.FC = () => {
  const [selectedNode, setSelectedNode] = useState<FlowNode | null>(null);
  const [highlightedPath, setHighlightedPath] = useState<string[]>([]);
  const [zoom, setZoom] = useState(1);
  const svgRef = useRef<SVGSVGElement>(null);

  // Define the flowchart nodes based on the PowerShell script
  const nodes: FlowNode[] = [
    {
      id: 'start',
      type: 'start',
      title: 'Script Start',
      description: 'Initialize DDG Automation',
      x: 400,
      y: 50,
      width: 120,
      height: 60,
      connections: ['params'],
      details: {
        purpose: 'Entry point for the Exchange Online Dynamic Distribution Group automation script',
        code: '# ddlcreate_new_exchange_online_runbook_enhanced.ps1'
      }
    },
    {
      id: 'params',
      type: 'process',
      title: 'Parameter Setup',
      description: 'Parse input parameters',
      x: 400,
      y: 150,
      width: 140,
      height: 60,
      connections: ['logging'],
      details: {
        purpose: 'Initialize script parameters with default values and validation',
        variables: ['$OrganizationDomain', '$WhatIf', '$LogPath', '$MaxRetries'],
        code: `param(
    [string]$OrganizationDomain = "yourtenant.onmicrosoft.com",
    [switch]$WhatIf = $false,
    [string]$LogPath = "",
    [int]$MaxRetries = 3
)`
      }
    },
    {
      id: 'logging',
      type: 'process',
      title: 'Initialize Logging',
      description: 'Setup transcript logging',
      x: 400,
      y: 250,
      width: 140,
      height: 60,
      connections: ['functions'],
      details: {
        purpose: 'Initialize logging system with timestamped output and optional file logging',
        functions: ['Write-TimestampedOutput'],
        code: `if ($LogPath) {
    $logFile = Join-Path $LogPath "DDG_Automation_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    Start-Transcript -Path $logFile -Append
}`
      }
    },
    {
      id: 'functions',
      type: 'function',
      title: 'Load Functions',
      description: 'Define helper functions',
      x: 400,
      y: 350,
      width: 140,
      height: 60,
      connections: ['connect'],
      details: {
        purpose: 'Load all helper functions for validation, parsing, and Exchange operations',
        functions: [
          'Test-DepartmentFormat',
          'Get-DepartmentComponents', 
          'New-RecipientFilter',
          'Invoke-ExchangeOperation'
        ],
        code: `function Write-TimestampedOutput { ... }
function Test-DepartmentFormat { ... }
function Get-DepartmentComponents { ... }`
      }
    },
    {
      id: 'connect',
      type: 'process',
      title: 'Connect Exchange',
      description: 'Connect to Exchange Online',
      x: 400,
      y: 450,
      width: 140,
      height: 60,
      connections: ['connect_check'],
      details: {
        purpose: 'Establish secure connection to Exchange Online using Managed Identity',
        code: 'Connect-ExchangeOnline -ManagedIdentity -Organization $OrganizationDomain',
        errorHandling: 'Retry logic with exponential backoff'
      }
    },
    {
      id: 'connect_check',
      type: 'decision',
      title: 'Connection Success?',
      description: 'Verify Exchange connection',
      x: 400,
      y: 550,
      width: 140,
      height: 80,
      connections: ['get_depts', 'connect_error'],
      details: {
        purpose: 'Validate Exchange Online connection before proceeding',
        errorHandling: 'Exit with code 1 if connection fails after max retries'
      }
    },
    {
      id: 'connect_error',
      type: 'error',
      title: 'Connection Failed',
      description: 'Exit with error',
      x: 600,
      y: 550,
      width: 120,
      height: 60,
      connections: ['end_error'],
      details: {
        purpose: 'Handle connection failure and exit gracefully',
        code: 'exit 1'
      }
    },
    {
      id: 'get_depts',
      type: 'process',
      title: 'Get Departments',
      description: 'Retrieve unique departments',
      x: 400,
      y: 670,
      width: 140,
      height: 60,
      connections: ['depts_check'],
      details: {
        purpose: 'Query Exchange to get all unique department names from recipients',
        code: `Get-Recipient -ResultSize Unlimited | 
    Where-Object {$_.Department -ne "" -and $_.RecipientTypeDetails -match "UserMailbox|MailUser"} | 
    Select-Object -Unique Department`,
        variables: ['$depts']
      }
    },
    {
      id: 'depts_check',
      type: 'decision',
      title: 'Departments Found?',
      description: 'Check if any departments exist',
      x: 400,
      y: 770,
      width: 140,
      height: 80,
      connections: ['init_stats', 'no_depts'],
      details: {
        purpose: 'Validate that departments were found before processing',
        code: 'if ($depts.Count -eq 0)'
      }
    },
    {
      id: 'no_depts',
      type: 'process',
      title: 'No Departments',
      description: 'Exit gracefully',
      x: 600,
      y: 770,
      width: 120,
      height: 60,
      connections: ['disconnect'],
      details: {
        purpose: 'Handle case where no valid departments are found',
        code: 'exit 0'
      }
    },
    {
      id: 'init_stats',
      type: 'process',
      title: 'Initialize Stats',
      description: 'Setup counters',
      x: 400,
      y: 890,
      width: 140,
      height: 60,
      connections: ['dept_loop'],
      details: {
        purpose: 'Initialize statistics tracking for reporting',
        variables: ['$stats.Created', '$stats.Updated', '$stats.Skipped', '$stats.ValidationErrors', '$stats.ProcessingErrors'],
        code: `$stats = @{
    Created = 0; Updated = 0; Skipped = 0
    ValidationErrors = 0; ProcessingErrors = 0
}`
      }
    },
    {
      id: 'dept_loop',
      type: 'process',
      title: 'For Each Department',
      description: 'Start department processing',
      x: 400,
      y: 990,
      width: 140,
      height: 60,
      connections: ['validate_format'],
      details: {
        purpose: 'Begin iterating through each department for processing',
        code: 'foreach ($dept in $depts)'
      }
    },
    {
      id: 'validate_format',
      type: 'function',
      title: 'Validate Format',
      description: 'Check department name format',
      x: 400,
      y: 1090,
      width: 140,
      height: 60,
      connections: ['format_check'],
      details: {
        purpose: 'Validate department name follows expected format: "12345 Name - USA"',
        functions: ['Test-DepartmentFormat'],
        code: `if ($DepartmentName -notmatch '^[0-9]{5}\\s+.+\\s+-\\s+[A-Z]{3}$')`
      }
    },
    {
      id: 'format_check',
      type: 'decision',
      title: 'Format Valid?',
      description: 'Check validation result',
      x: 400,
      y: 1190,
      width: 140,
      height: 80,
      connections: ['parse_components', 'format_error'],
      details: {
        purpose: 'Determine if department name format is valid before parsing',
        errorHandling: 'Skip department and increment validation error counter'
      }
    },
    {
      id: 'format_error',
      type: 'error',
      title: 'Format Error',
      description: 'Skip department',
      x: 200,
      y: 1190,
      width: 120,
      height: 60,
      connections: ['next_dept'],
      details: {
        purpose: 'Handle invalid department format and continue to next department',
        code: '$stats.ValidationErrors++'
      }
    },
    {
      id: 'parse_components',
      type: 'function',
      title: 'Parse Components',
      description: 'Extract dept number, name, country',
      x: 400,
      y: 1310,
      width: 140,
      height: 60,
      connections: ['create_filter'],
      details: {
        purpose: 'Extract department number, name, and country code from department string',
        functions: ['Get-DepartmentComponents'],
        variables: ['$components.Number', '$components.Name', '$components.CountryCode'],
        code: `$deptNum = $DepartmentName.Substring(0, 5)
$countryCode = $DepartmentName.Substring($DepartmentName.Length - 3, 3)
$deptName = $DepartmentName.Substring($startPos, $nameLength)`
      }
    },
    {
      id: 'create_filter',
      type: 'function',
      title: 'Create Filter',
      description: 'Build recipient filter',
      x: 400,
      y: 1430,
      width: 140,
      height: 60,
      connections: ['whatif_check'],
      details: {
        purpose: 'Create OPATH recipient filter for the Dynamic Distribution Group',
        functions: ['New-RecipientFilter'],
        code: `"((RecipientTypeDetails -eq 'UserMailbox' -or RecipientTypeDetails -eq 'MailUser') " +
"-and Department -like '$DepartmentNumber*' " +
"-and Department -like '*$CountryCode')"`,
        variables: ['$recipientFilter']
      }
    },
    {
      id: 'whatif_check',
      type: 'decision',
      title: 'WhatIf Mode?',
      description: 'Check if in test mode',
      x: 400,
      y: 1530,
      width: 140,
      height: 80,
      connections: ['whatif_output', 'check_existing'],
      details: {
        purpose: 'Check if script is running in WhatIf (test) mode',
        code: 'if ($WhatIf)'
      }
    },
    {
      id: 'whatif_output',
      type: 'process',
      title: 'WhatIf Output',
      description: 'Log what would happen',
      x: 200,
      y: 1530,
      width: 120,
      height: 60,
      connections: ['next_dept'],
      details: {
        purpose: 'Output what would happen without making changes',
        code: 'Write-TimestampedOutput "WHAT-IF: Would process group $groupName"'
      }
    },
    {
      id: 'check_existing',
      type: 'process',
      title: 'Check Existing',
      description: 'Look for existing recipient',
      x: 400,
      y: 1650,
      width: 140,
      height: 60,
      connections: ['existing_check'],
      details: {
        purpose: 'Check if any recipient already exists with the target group name',
        code: 'Get-Recipient -Identity $groupName -ErrorAction SilentlyContinue',
        functions: ['Invoke-ExchangeOperation']
      }
    },
    {
      id: 'existing_check',
      type: 'decision',
      title: 'Recipient Exists?',
      description: 'Check if name is taken',
      x: 400,
      y: 1750,
      width: 140,
      height: 80,
      connections: ['check_ddg_type', 'create_new'],
      details: {
        purpose: 'Determine if the target group name is already in use',
        code: 'if ($existingRecipient)'
      }
    },
    {
      id: 'check_ddg_type',
      type: 'process',
      title: 'Check DDG Type',
      description: 'Verify if it\'s a DDG',
      x: 600,
      y: 1750,
      width: 140,
      height: 60,
      connections: ['ddg_type_check'],
      details: {
        purpose: 'Check if existing recipient is specifically a Dynamic Distribution Group',
        code: 'Get-DynamicDistributionGroup -Identity $groupName -ErrorAction SilentlyContinue'
      }
    },
    {
      id: 'ddg_type_check',
      type: 'decision',
      title: 'Is DDG?',
      description: 'Check recipient type',
      x: 600,
      y: 1850,
      width: 140,
      height: 80,
      connections: ['update_ddg', 'skip_conflict'],
      details: {
        purpose: 'Determine if existing recipient is a DDG that can be updated',
        code: 'if ($existingDDG)'
      }
    },
    {
      id: 'update_ddg',
      type: 'process',
      title: 'Update DDG',
      description: 'Update existing group',
      x: 600,
      y: 1970,
      width: 140,
      height: 60,
      connections: ['update_success'],
      details: {
        purpose: 'Update existing Dynamic Distribution Group with new settings',
        code: `Set-DynamicDistributionGroup -Identity $groupName 
    -DisplayName $displayName 
    -RecipientFilter $recipientFilter`,
        functions: ['Invoke-ExchangeOperation']
      }
    },
    {
      id: 'update_success',
      type: 'success',
      title: 'Updated',
      description: 'DDG updated successfully',
      x: 600,
      y: 2090,
      width: 120,
      height: 60,
      connections: ['next_dept'],
      details: {
        purpose: 'Record successful DDG update',
        code: '$stats.Updated++'
      }
    },
    {
      id: 'skip_conflict',
      type: 'process',
      title: 'Skip Conflict',
      description: 'Name conflict with other type',
      x: 800,
      y: 1850,
      width: 140,
      height: 60,
      connections: ['next_dept'],
      details: {
        purpose: 'Skip creation due to name conflict with different recipient type',
        code: '$stats.Skipped++',
        errorHandling: 'Log conflict and continue processing'
      }
    },
    {
      id: 'create_new',
      type: 'process',
      title: 'Create New DDG',
      description: 'Create new group',
      x: 400,
      y: 1870,
      width: 140,
      height: 60,
      connections: ['create_success'],
      details: {
        purpose: 'Create new Dynamic Distribution Group with specified settings',
        code: `New-DynamicDistributionGroup -Name $groupName 
    -Alias $groupName 
    -DisplayName $displayName 
    -RecipientFilter $recipientFilter`,
        functions: ['Invoke-ExchangeOperation']
      }
    },
    {
      id: 'create_success',
      type: 'success',
      title: 'Created',
      description: 'DDG created successfully',
      x: 400,
      y: 1990,
      width: 120,
      height: 60,
      connections: ['next_dept'],
      details: {
        purpose: 'Record successful DDG creation',
        code: '$stats.Created++'
      }
    },
    {
      id: 'next_dept',
      type: 'process',
      title: 'Next Department',
      description: 'Continue to next department',
      x: 100,
      y: 1500,
      width: 140,
      height: 60,
      connections: ['loop_check'],
      details: {
        purpose: 'Move to the next department in the processing loop',
        code: 'continue'
      }
    },
    {
      id: 'loop_check',
      type: 'decision',
      title: 'More Departments?',
      description: 'Check if loop continues',
      x: 100,
      y: 1400,
      width: 140,
      height: 80,
      connections: ['validate_format', 'generate_report'],
      details: {
        purpose: 'Check if there are more departments to process',
        code: 'foreach loop condition'
      }
    },
    {
      id: 'generate_report',
      type: 'process',
      title: 'Generate Report',
      description: 'Create summary report',
      x: 100,
      y: 1200,
      width: 140,
      height: 60,
      connections: ['disconnect'],
      details: {
        purpose: 'Generate comprehensive summary report with statistics',
        variables: ['$successCount', '$totalErrors', '$successRate', '$errorRate'],
        code: `Write-TimestampedOutput "✓ Created: $($stats.Created)"
Write-TimestampedOutput "✓ Updated: $($stats.Updated)"
Write-TimestampedOutput "⚠ Skipped: $($stats.Skipped)"`
      }
    },
    {
      id: 'disconnect',
      type: 'process',
      title: 'Disconnect',
      description: 'Close Exchange connection',
      x: 100,
      y: 1000,
      width: 140,
      height: 60,
      connections: ['end_success'],
      details: {
        purpose: 'Cleanly disconnect from Exchange Online',
        code: 'Disconnect-ExchangeOnline -Confirm:$false'
      }
    },
    {
      id: 'end_success',
      type: 'end',
      title: 'End Success',
      description: 'Script completed successfully',
      x: 100,
      y: 800,
      width: 120,
      height: 60,
      connections: [],
      details: {
        purpose: 'Script execution completed successfully',
        code: 'exit 0'
      }
    },
    {
      id: 'end_error',
      type: 'end',
      title: 'End Error',
      description: 'Script completed with errors',
      x: 600,
      y: 450,
      width: 120,
      height: 60,
      connections: [],
      details: {
        purpose: 'Script execution completed with errors',
        code: 'exit 1'
      }
    }
  ];

  const getNodeColor = (type: string) => {
    const colors = {
      start: '#22c55e',
      end: '#ef4444',
      process: '#3b82f6',
      decision: '#f59e0b',
      function: '#8b5cf6',
      error: '#ef4444',
      success: '#22c55e'
    };
    return colors[type as keyof typeof colors] || '#6b7280';
  };

  const getNodeIcon = (type: string) => {
    const icons = {
      start: Play,
      end: Square,
      process: Circle,
      decision: Diamond,
      function: Settings,
      error: XCircle,
      success: CheckCircle
    };
    const IconComponent = icons[type as keyof typeof icons] || Circle;
    return <IconComponent className="w-4 h-4" />;
  };

  const handleNodeClick = (node: FlowNode) => {
    setSelectedNode(node);
    // Highlight the path from this node
    const path = [node.id, ...node.connections];
    setHighlightedPath(path);
  };

  const handleZoomIn = () => setZoom(prev => Math.min(prev + 0.2, 3));
  const handleZoomOut = () => setZoom(prev => Math.max(prev - 0.2, 0.5));
  const handleReset = () => {
    setZoom(1);
    setHighlightedPath([]);
    setSelectedNode(null);
  };

  const renderNode = (node: FlowNode) => {
    const isHighlighted = highlightedPath.includes(node.id);
    const isSelected = selectedNode?.id === node.id;
    
    return (
      <g key={node.id} transform={`translate(${node.x}, ${node.y})`}>
        {/* Node background */}
        <rect
          x={0}
          y={0}
          width={node.width}
          height={node.height}
          rx={node.type === 'decision' ? 0 : 8}
          fill={isSelected ? '#1e40af' : getNodeColor(node.type)}
          stroke={isHighlighted ? '#fbbf24' : '#e5e7eb'}
          strokeWidth={isHighlighted ? 3 : 1}
          className="cursor-pointer transition-all duration-200 hover:stroke-yellow-400 hover:stroke-2"
          onClick={() => handleNodeClick(node)}
          transform={node.type === 'decision' ? `rotate(45 ${node.width/2} ${node.height/2})` : ''}
        />
        
        {/* Node content */}
        <g className="cursor-pointer" onClick={() => handleNodeClick(node)}>
          <text
            x={node.width / 2}
            y={node.height / 2 - 8}
            textAnchor="middle"
            className="fill-white text-xs font-semibold"
            style={{ fontSize: '11px' }}
          >
            {node.title}
          </text>
          <text
            x={node.width / 2}
            y={node.height / 2 + 8}
            textAnchor="middle"
            className="fill-white text-xs opacity-90"
            style={{ fontSize: '9px' }}
          >
            {node.description}
          </text>
        </g>
      </g>
    );
  };

  const renderConnections = () => {
    return nodes.flatMap(node => 
      node.connections.map(targetId => {
        const target = nodes.find(n => n.id === targetId);
        if (!target) return null;

        const isHighlighted = highlightedPath.includes(node.id) && highlightedPath.includes(targetId);
        
        return (
          <line
            key={`${node.id}-${targetId}`}
            x1={node.x + node.width / 2}
            y1={node.y + node.height}
            x2={target.x + target.width / 2}
            y2={target.y}
            stroke={isHighlighted ? '#fbbf24' : '#6b7280'}
            strokeWidth={isHighlighted ? 3 : 2}
            markerEnd="url(#arrowhead)"
            className="transition-all duration-200"
          />
        );
      })
    ).filter(Boolean);
  };

  return (
    <div className="max-w-7xl mx-auto p-6 space-y-6">
      <div className="text-center space-y-4">
        <h1 className="text-4xl font-bold text-gray-900">
          Exchange Online DDG Script Flowchart
        </h1>
        <p className="text-xl text-gray-600 max-w-4xl mx-auto">
          Interactive visualization of the PowerShell automation script workflow with step-by-step process flow
        </p>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Flowchart Canvas */}
        <div className="lg:col-span-2">
          <Card className="h-[800px]">
            <CardHeader className="pb-2">
              <div className="flex items-center justify-between">
                <CardTitle className="flex items-center space-x-2">
                  <GitBranch className="w-5 h-5" />
                  <span>Script Flow Diagram</span>
                </CardTitle>
                <div className="flex items-center space-x-2">
                  <Button variant="outline" size="sm" onClick={handleZoomOut}>
                    <ZoomOut className="w-4 h-4" />
                  </Button>
                  <Badge variant="outline">{Math.round(zoom * 100)}%</Badge>
                  <Button variant="outline" size="sm" onClick={handleZoomIn}>
                    <ZoomIn className="w-4 h-4" />
                  </Button>
                  <Button variant="outline" size="sm" onClick={handleReset}>
                    <RotateCcw className="w-4 h-4" />
                  </Button>
                </div>
              </div>
            </CardHeader>
            <CardContent className="p-0">
              <div className="relative w-full h-[720px] overflow-auto bg-gray-50">
                <svg
                  ref={svgRef}
                  width="1000"
                  height="2200"
                  className="w-full h-full"
                  style={{ transform: `scale(${zoom})` }}
                >
                  <defs>
                    <marker
                      id="arrowhead"
                      markerWidth="10"
                      markerHeight="7"
                      refX="9"
                      refY="3.5"
                      orient="auto"
                    >
                      <polygon
                        points="0 0, 10 3.5, 0 7"
                        fill="#6b7280"
                      />
                    </marker>
                  </defs>
                  
                  {renderConnections()}
                  {nodes.map(renderNode)}
                </svg>
              </div>
            </CardContent>
          </Card>
        </div>

        {/* Details Panel */}
        <div className="space-y-4">
          {/* Legend */}
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center space-x-2">
                <Target className="w-5 h-5" />
                <span>Legend</span>
              </CardTitle>
            </CardHeader>
            <CardContent className="space-y-3">
              <div className="grid grid-cols-2 gap-2 text-sm">
                <div className="flex items-center space-x-2">
                  <div className="w-4 h-4 bg-green-500 rounded"></div>
                  <span>Start/Success</span>
                </div>
                <div className="flex items-center space-x-2">
                  <div className="w-4 h-4 bg-red-500 rounded"></div>
                  <span>End/Error</span>
                </div>
                <div className="flex items-center space-x-2">
                  <div className="w-4 h-4 bg-blue-500 rounded"></div>
                  <span>Process</span>
                </div>
                <div className="flex items-center space-x-2">
                  <div className="w-4 h-4 bg-yellow-500 rounded transform rotate-45"></div>
                  <span>Decision</span>
                </div>
                <div className="flex items-center space-x-2">
                  <div className="w-4 h-4 bg-purple-500 rounded"></div>
                  <span>Function</span>
                </div>
                <div className="flex items-center space-x-2">
                  <div className="w-3 h-0.5 bg-yellow-400"></div>
                  <span>Highlighted</span>
                </div>
              </div>
            </CardContent>
          </Card>

          {/* Node Details */}
          {selectedNode ? (
            <Card>
              <CardHeader>
                <CardTitle className="flex items-center space-x-2">
                  {getNodeIcon(selectedNode.type)}
                  <span>{selectedNode.title}</span>
                </CardTitle>
                <CardDescription>{selectedNode.description}</CardDescription>
              </CardHeader>
              <CardContent className="space-y-4">
                <div>
                  <h4 className="font-semibold text-sm mb-2">Purpose:</h4>
                  <p className="text-sm text-gray-600">{selectedNode.details.purpose}</p>
                </div>

                {selectedNode.details.code && (
                  <div>
                    <h4 className="font-semibold text-sm mb-2 flex items-center space-x-1">
                      <Code className="w-4 h-4" />
                      <span>Code:</span>
                    </h4>
                    <pre className="bg-gray-900 text-gray-100 p-3 rounded text-xs overflow-x-auto">
                      <code>{selectedNode.details.code}</code>
                    </pre>
                  </div>
                )}

                {selectedNode.details.variables && (
                  <div>
                    <h4 className="font-semibold text-sm mb-2 flex items-center space-x-1">
                      <Database className="w-4 h-4" />
                      <span>Variables:</span>
                    </h4>
                    <div className="flex flex-wrap gap-1">
                      {selectedNode.details.variables.map((variable, idx) => (
                        <Badge key={idx} variant="outline" className="text-xs">
                          {variable}
                        </Badge>
                      ))}
                    </div>
                  </div>
                )}

                {selectedNode.details.functions && (
                  <div>
                    <h4 className="font-semibold text-sm mb-2 flex items-center space-x-1">
                      <Settings className="w-4 h-4" />
                      <span>Functions:</span>
                    </h4>
                    <div className="flex flex-wrap gap-1">
                      {selectedNode.details.functions.map((func, idx) => (
                        <Badge key={idx} variant="secondary" className="text-xs">
                          {func}
                        </Badge>
                      ))}
                    </div>
                  </div>
                )}

                {selectedNode.details.errorHandling && (
                  <div>
                    <h4 className="font-semibold text-sm mb-2 flex items-center space-x-1">
                      <Shield className="w-4 h-4" />
                      <span>Error Handling:</span>
                    </h4>
                    <p className="text-sm text-gray-600">{selectedNode.details.errorHandling}</p>
                  </div>
                )}
              </CardContent>
            </Card>
          ) : (
            <Card>
              <CardHeader>
                <CardTitle className="flex items-center space-x-2">
                  <Target className="w-5 h-5" />
                  <span>Node Details</span>
                </CardTitle>
              </CardHeader>
              <CardContent>
                <p className="text-sm text-gray-600">
                  Click on any node in the flowchart to see detailed information about that step, 
                  including code snippets, variables, functions, and error handling.
                </p>
              </CardContent>
            </Card>
          )}

          {/* Script Statistics */}
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center space-x-2">
                <FileText className="w-5 h-5" />
                <span>Script Overview</span>
              </CardTitle>
            </CardHeader>
            <CardContent className="space-y-3">
              <div className="grid grid-cols-2 gap-4 text-sm">
                <div>
                  <div className="font-semibold">Total Nodes</div>
                  <div className="text-2xl font-bold text-blue-600">{nodes.length}</div>
                </div>
                <div>
                  <div className="font-semibold">Decision Points</div>
                  <div className="text-2xl font-bold text-yellow-600">
                    {nodes.filter(n => n.type === 'decision').length}
                  </div>
                </div>
                <div>
                  <div className="font-semibold">Functions</div>
                  <div className="text-2xl font-bold text-purple-600">
                    {nodes.filter(n => n.type === 'function').length}
                  </div>
                </div>
                <div>
                  <div className="font-semibold">Error Paths</div>
                  <div className="text-2xl font-bold text-red-600">
                    {nodes.filter(n => n.type === 'error').length}
                  </div>
                </div>
              </div>
              
              <Separator />
              
              <div className="space-y-2">
                <h4 className="font-semibold text-sm">Key Features:</h4>
                <ul className="text-xs space-y-1 text-gray-600">
                  <li>• Comprehensive error handling with retry logic</li>
                  <li>• Department format validation and parsing</li>
                  <li>• Conflict detection and resolution</li>
                  <li>• WhatIf mode for safe testing</li>
                  <li>• Detailed logging and reporting</li>
                  <li>• Managed Identity authentication</li>
                </ul>
              </div>
            </CardContent>
          </Card>
        </div>
      </div>
    </div>
  );
};

export default InteractiveFlowchart;