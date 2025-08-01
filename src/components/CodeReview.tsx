import React, { useState } from 'react';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from './ui/card';
import { Badge } from './ui/badge';
import { Tabs, TabsContent, TabsList, TabsTrigger } from './ui/tabs';
import { Alert, AlertDescription } from './ui/alert';
import { Button } from './ui/button';
import { Collapsible, CollapsibleContent, CollapsibleTrigger } from './ui/collapsible';
import { 
  CheckCircle, 
  XCircle, 
  AlertTriangle, 
  Info, 
  ChevronDown, 
  ChevronRight,
  Code,
  Shield,
  Zap,
  Settings,
  FileText,
  Target,
  Clock,
  Database
} from 'lucide-react';

interface ReviewItem {
  category: string;
  title: string;
  status: 'excellent' | 'good' | 'warning' | 'critical';
  description: string;
  details: string[];
  codeExample?: string;
  recommendation?: string;
}

const CodeReview: React.FC = () => {
  const [expandedSections, setExpandedSections] = useState<string[]>(['overview']);

  const toggleSection = (section: string) => {
    setExpandedSections(prev => 
      prev.includes(section) 
        ? prev.filter(s => s !== section)
        : [...prev, section]
    );
  };

  const reviewItems: ReviewItem[] = [
    {
      category: 'Architecture',
      title: 'Script Structure & Organization',
      status: 'excellent',
      description: 'Well-structured with proper function separation and clear execution flow',
      details: [
        'Proper parameter handling with validation',
        'Modular functions for reusability',
        'Clear separation of concerns',
        'Comprehensive error handling throughout'
      ],
      codeExample: `param(
    [Parameter(Mandatory=$false)]
    [string]$OrganizationDomain = "yourtenant.onmicrosoft.com",
    [Parameter(Mandatory=$false)]
    [switch]$WhatIf = $false,
    [Parameter(Mandatory=$false)]
    [string]$LogPath = "",
    [Parameter(Mandatory=$false)]
    [int]$MaxRetries = 3
)`
    },
    {
      category: 'Security',
      title: 'Authentication & Security Practices',
      status: 'excellent',
      description: 'Implements secure authentication with managed identity',
      details: [
        'Uses Managed Identity for secure authentication',
        'No hardcoded credentials or secrets',
        'Proper connection cleanup on exit',
        'Secure parameter handling'
      ],
      codeExample: `Connect-ExchangeOnline -ManagedIdentity -Organization $OrganizationDomain`
    },
    {
      category: 'Error Handling',
      title: 'Comprehensive Error Management',
      status: 'excellent',
      description: 'Robust error handling with retry logic and detailed reporting',
      details: [
        'Retry logic with exponential backoff',
        'Comprehensive error categorization',
        'Graceful failure handling',
        'Detailed error reporting and logging'
      ],
      codeExample: `function Invoke-ExchangeOperation {
    param(
        [scriptblock]$Operation,
        [string]$OperationName,
        [int]$MaxRetries = 3
    )
    
    $attempt = 1
    while ($attempt -le $MaxRetries) {
        try {
            $result = & $Operation
            return @{ Success = $true; Result = $result; Error = $null }
        }
        catch {
            if ($attempt -eq $MaxRetries) {
                return @{ Success = $false; Result = $null; Error = $_.Exception.Message }
            }
            $attempt++
            Start-Sleep -Seconds (2 * $attempt)  # Exponential backoff
        }
    }
}`
    },
    {
      category: 'Validation',
      title: 'Department Name Validation',
      status: 'excellent',
      description: 'Comprehensive validation of department name format and components',
      details: [
        'Regex-based format validation',
        'Component extraction with bounds checking',
        'Proper error messaging for validation failures',
        'Safe string manipulation with length checks'
      ],
      codeExample: `function Test-DepartmentFormat {
    param([string]$DepartmentName)
    
    if ($DepartmentName.Length -lt 13) {
        return @{
            IsValid = $false
            Error = "Department name too short (minimum 13 characters required)"
        }
    }
    
    if ($DepartmentName -notmatch '^[0-9]{5}\\s+.+\\s+-\\s+[A-Z]{3}$') {
        return @{
            IsValid = $false
            Error = "Invalid format. Expected: '12345 Department Name - USA'"
        }
    }
    
    return @{ IsValid = $true; Error = $null }
}`
    },
    {
      category: 'Performance',
      title: 'Query Optimization',
      status: 'good',
      description: 'Efficient recipient filtering with room for minor improvements',
      details: [
        'Single query to get all departments',
        'Proper filtering of recipient types',
        'Efficient OPATH filter construction',
        'Could benefit from batch processing for large datasets'
      ],
      codeExample: `$recipientFilter = "((RecipientTypeDetails -eq 'UserMailbox' -or RecipientTypeDetails -eq 'MailUser') " +
                   "-and Department -like '$DepartmentNumber*' " +
                   "-and Department -like '*$CountryCode' " +
                   "-and Name -notlike 'SystemMailbox*' " +
                   "-and RecipientTypeDetails -ne 'SharedMailbox')"`,
      recommendation: 'Consider implementing batch processing for organizations with >1000 departments'
    },
    {
      category: 'Logging',
      title: 'Comprehensive Logging & Reporting',
      status: 'excellent',
      description: 'Detailed logging with timestamps and comprehensive reporting',
      details: [
        'Timestamped logging function',
        'Optional transcript logging to file',
        'Detailed statistics and success rates',
        'Comprehensive error reporting',
        'Clear execution summary'
      ],
      codeExample: `function Write-TimestampedOutput {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $output = "[$timestamp] $Message"
    Write-Output $output
}`
    },
    {
      category: 'Reliability',
      title: 'Conflict Detection & Resolution',
      status: 'excellent',
      description: 'Proper handling of existing recipients and group conflicts',
      details: [
        'Checks for existing recipients before creation',
        'Distinguishes between DDGs and other recipient types',
        'Updates existing DDGs instead of failing',
        'Skips conflicting recipient types safely'
      ],
      codeExample: `$existingRecipient = Get-Recipient -Identity $groupName -ErrorAction SilentlyContinue

if ($existingRecipient) {
    $existingDDG = Get-DynamicDistributionGroup -Identity $groupName -ErrorAction SilentlyContinue
    
    if ($existingDDG) {
        # Update existing DDG
        Set-DynamicDistributionGroup -Identity $groupName -DisplayName $displayName
    } else {
        # Skip - exists as different recipient type
        Write-Output "SKIPPING: $groupName already exists as $($existingRecipient.RecipientTypeDetails)"
    }
}`
    },
    {
      category: 'Configuration',
      title: 'Recipient Filter Configuration',
      status: 'good',
      description: 'Comprehensive filtering with enhanced exclusions',
      details: [
        'Excludes system mailboxes properly',
        'Filters out service accounts',
        'Includes both UserMailbox and MailUser types',
        'Enhanced exclusions for health/discovery mailboxes'
      ],
      recommendation: 'Consider adding organization-specific exclusion patterns as parameters'
    },
    {
      category: 'Maintainability',
      title: 'Code Maintainability & Documentation',
      status: 'excellent',
      description: 'Well-documented with clear function names and comprehensive comments',
      details: [
        'Clear function names and parameter documentation',
        'Comprehensive inline comments',
        'Modular design for easy modification',
        'Consistent coding style throughout'
      ]
    },
    {
      category: 'Testing',
      title: 'Testing & Validation Features',
      status: 'excellent',
      description: 'Includes WhatIf parameter for safe testing',
      details: [
        'WhatIf parameter for dry-run testing',
        'Comprehensive validation before execution',
        'Detailed logging for troubleshooting',
        'Exit codes for automation integration'
      ],
      codeExample: `if ($WhatIf) {
    Write-TimestampedOutput "WHAT-IF: Would process group $groupName"
    continue
}`
    }
  ];

  const getStatusIcon = (status: string) => {
    switch (status) {
      case 'excellent': return <CheckCircle className="h-5 w-5 text-green-500" />;
      case 'good': return <CheckCircle className="h-5 w-5 text-blue-500" />;
      case 'warning': return <AlertTriangle className="h-5 w-5 text-yellow-500" />;
      case 'critical': return <XCircle className="h-5 w-5 text-red-500" />;
      default: return <Info className="h-5 w-5 text-gray-500" />;
    }
  };

  const getStatusBadge = (status: string) => {
    const variants = {
      excellent: 'bg-green-100 text-green-800 border-green-200',
      good: 'bg-blue-100 text-blue-800 border-blue-200',
      warning: 'bg-yellow-100 text-yellow-800 border-yellow-200',
      critical: 'bg-red-100 text-red-800 border-red-200'
    };
    
    return (
      <Badge className={variants[status as keyof typeof variants] || 'bg-gray-100 text-gray-800'}>
        {status.toUpperCase()}
      </Badge>
    );
  };

  const getCategoryIcon = (category: string) => {
    const icons = {
      'Architecture': <Settings className="h-4 w-4" />,
      'Security': <Shield className="h-4 w-4" />,
      'Error Handling': <AlertTriangle className="h-4 w-4" />,
      'Validation': <CheckCircle className="h-4 w-4" />,
      'Performance': <Zap className="h-4 w-4" />,
      'Logging': <FileText className="h-4 w-4" />,
      'Reliability': <Target className="h-4 w-4" />,
      'Configuration': <Database className="h-4 w-4" />,
      'Maintainability': <Code className="h-4 w-4" />,
      'Testing': <Clock className="h-4 w-4" />
    };
    return icons[category as keyof typeof icons] || <Info className="h-4 w-4" />;
  };

  const overallScore = reviewItems.reduce((acc, item) => {
    const scores = { excellent: 4, good: 3, warning: 2, critical: 1 };
    return acc + scores[item.status];
  }, 0);
  const maxScore = reviewItems.length * 4;
  const scorePercentage = Math.round((overallScore / maxScore) * 100);

  const statusCounts = reviewItems.reduce((acc, item) => {
    acc[item.status] = (acc[item.status] || 0) + 1;
    return acc;
  }, {} as Record<string, number>);

  return (
    <div className="max-w-6xl mx-auto p-6 space-y-6">
      <div className="text-center space-y-4">
        <h1 className="text-4xl font-bold text-gray-900">
          Exchange Online DDG Script Code Review
        </h1>
        <p className="text-xl text-gray-600 max-w-3xl mx-auto">
          Comprehensive analysis of the PowerShell automation script for Dynamic Distribution Group management
        </p>
      </div>

      {/* Overall Score Card */}
      <Card className="border-2 border-blue-200 bg-gradient-to-r from-blue-50 to-indigo-50">
        <CardHeader className="text-center">
          <CardTitle className="text-2xl text-blue-900">Overall Code Quality Score</CardTitle>
          <div className="text-6xl font-bold text-blue-600 my-4">{scorePercentage}%</div>
          <CardDescription className="text-lg">
            Excellent code quality with robust error handling and comprehensive validation
          </CardDescription>
        </CardHeader>
        <CardContent>
          <div className="grid grid-cols-2 md:grid-cols-4 gap-4 text-center">
            <div className="space-y-2">
              <div className="text-2xl font-bold text-green-600">{statusCounts.excellent || 0}</div>
              <div className="text-sm text-gray-600">Excellent</div>
            </div>
            <div className="space-y-2">
              <div className="text-2xl font-bold text-blue-600">{statusCounts.good || 0}</div>
              <div className="text-sm text-gray-600">Good</div>
            </div>
            <div className="space-y-2">
              <div className="text-2xl font-bold text-yellow-600">{statusCounts.warning || 0}</div>
              <div className="text-sm text-gray-600">Warnings</div>
            </div>
            <div className="space-y-2">
              <div className="text-2xl font-bold text-red-600">{statusCounts.critical || 0}</div>
              <div className="text-sm text-gray-600">Critical</div>
            </div>
          </div>
        </CardContent>
      </Card>

      <Tabs defaultValue="detailed" className="w-full">
        <TabsList className="grid w-full grid-cols-3">
          <TabsTrigger value="detailed">Detailed Review</TabsTrigger>
          <TabsTrigger value="summary">Executive Summary</TabsTrigger>
          <TabsTrigger value="recommendations">Recommendations</TabsTrigger>
        </TabsList>

        <TabsContent value="detailed" className="space-y-4">
          {reviewItems.map((item, index) => (
            <Card key={index} className="border-l-4 border-l-blue-500">
              <Collapsible 
                open={expandedSections.includes(`item-${index}`)}
                onOpenChange={() => toggleSection(`item-${index}`)}
              >
                <CollapsibleTrigger asChild>
                  <CardHeader className="cursor-pointer hover:bg-gray-50 transition-colors">
                    <div className="flex items-center justify-between">
                      <div className="flex items-center space-x-3">
                        {getCategoryIcon(item.category)}
                        <div>
                          <CardTitle className="text-lg">{item.title}</CardTitle>
                          <CardDescription className="flex items-center space-x-2 mt-1">
                            <Badge variant="outline" className="text-xs">
                              {item.category}
                            </Badge>
                            {getStatusBadge(item.status)}
                          </CardDescription>
                        </div>
                      </div>
                      <div className="flex items-center space-x-2">
                        {getStatusIcon(item.status)}
                        {expandedSections.includes(`item-${index}`) ? 
                          <ChevronDown className="h-4 w-4" /> : 
                          <ChevronRight className="h-4 w-4" />
                        }
                      </div>
                    </div>
                  </CardHeader>
                </CollapsibleTrigger>
                
                <CollapsibleContent>
                  <CardContent className="pt-0">
                    <p className="text-gray-700 mb-4">{item.description}</p>
                    
                    <div className="space-y-4">
                      <div>
                        <h4 className="font-semibold text-gray-900 mb-2">Key Points:</h4>
                        <ul className="list-disc list-inside space-y-1 text-gray-700">
                          {item.details.map((detail, idx) => (
                            <li key={idx}>{detail}</li>
                          ))}
                        </ul>
                      </div>

                      {item.codeExample && (
                        <div>
                          <h4 className="font-semibold text-gray-900 mb-2">Code Example:</h4>
                          <pre className="bg-gray-900 text-gray-100 p-4 rounded-lg overflow-x-auto text-sm">
                            <code>{item.codeExample}</code>
                          </pre>
                        </div>
                      )}

                      {item.recommendation && (
                        <Alert>
                          <Info className="h-4 w-4" />
                          <AlertDescription>
                            <strong>Recommendation:</strong> {item.recommendation}
                          </AlertDescription>
                        </Alert>
                      )}
                    </div>
                  </CardContent>
                </CollapsibleContent>
              </Collapsible>
            </Card>
          ))}
        </TabsContent>

        <TabsContent value="summary" className="space-y-6">
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center space-x-2">
                <CheckCircle className="h-6 w-6 text-green-500" />
                <span>Executive Summary</span>
              </CardTitle>
            </CardHeader>
            <CardContent className="space-y-4">
              <div className="prose max-w-none">
                <h3 className="text-lg font-semibold text-gray-900">Script Assessment</h3>
                <p className="text-gray-700">
                  The Exchange Online Dynamic Distribution Group automation script demonstrates <strong>excellent code quality</strong> 
                  with a score of {scorePercentage}%. The script is well-architected, secure, and production-ready.
                </p>

                <h3 className="text-lg font-semibold text-gray-900 mt-6">Key Strengths</h3>
                <ul className="list-disc list-inside space-y-2 text-gray-700">
                  <li><strong>Security:</strong> Uses Managed Identity authentication with no hardcoded credentials</li>
                  <li><strong>Reliability:</strong> Comprehensive error handling with retry logic and exponential backoff</li>
                  <li><strong>Validation:</strong> Robust department name format validation and component parsing</li>
                  <li><strong>Conflict Resolution:</strong> Intelligent handling of existing recipients and groups</li>
                  <li><strong>Logging:</strong> Detailed timestamped logging with comprehensive reporting</li>
                  <li><strong>Testing:</strong> WhatIf parameter for safe dry-run testing</li>
                </ul>

                <h3 className="text-lg font-semibold text-gray-900 mt-6">Production Readiness</h3>
                <p className="text-gray-700">
                  This script is <strong>production-ready</strong> and follows PowerShell best practices. It includes:
                </p>
                <ul className="list-disc list-inside space-y-1 text-gray-700">
                  <li>Proper parameter validation and error handling</li>
                  <li>Comprehensive logging and audit trails</li>
                  <li>Safe execution with conflict detection</li>
                  <li>Modular design for maintainability</li>
                  <li>Exit codes for automation integration</li>
                </ul>
              </div>
            </CardContent>
          </Card>
        </TabsContent>

        <TabsContent value="recommendations" className="space-y-4">
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center space-x-2">
                <Target className="h-6 w-6 text-blue-500" />
                <span>Optimization Recommendations</span>
              </CardTitle>
              <CardDescription>
                Suggestions for further enhancement and optimization
              </CardDescription>
            </CardHeader>
            <CardContent className="space-y-6">
              <div className="space-y-4">
                <Alert>
                  <Zap className="h-4 w-4" />
                  <AlertDescription>
                    <strong>Performance Enhancement:</strong> For organizations with &gt;1000 departments, 
                    consider implementing batch processing to reduce memory usage and improve execution time.
                  </AlertDescription>
                </Alert>

                <Alert>
                  <Settings className="h-4 w-4" />
                  <AlertDescription>
                    <strong>Configuration Flexibility:</strong> Add organization-specific exclusion patterns 
                    as parameters to handle custom mailbox naming conventions.
                  </AlertDescription>
                </Alert>

                <Alert>
                  <Database className="h-4 w-4" />
                  <AlertDescription>
                    <strong>Monitoring Integration:</strong> Consider adding integration with monitoring 
                    systems for automated alerting on script failures or high error rates.
                  </AlertDescription>
                </Alert>

                <Alert>
                  <Clock className="h-4 w-4" />
                  <AlertDescription>
                    <strong>Scheduling Optimization:</strong> Implement intelligent scheduling to avoid 
                    peak usage times and reduce impact on Exchange Online performance.
                  </AlertDescription>
                </Alert>
              </div>

              <div className="bg-green-50 border border-green-200 rounded-lg p-4">
                <h4 className="font-semibold text-green-900 mb-2">Overall Assessment</h4>
                <p className="text-green-800">
                  The script is <strong>excellently configured and optimized</strong> for its intended purpose. 
                  It demonstrates professional-grade PowerShell development with comprehensive error handling, 
                  security best practices, and production-ready features. The suggested enhancements are 
                  minor optimizations rather than critical fixes.
                </p>
              </div>
            </CardContent>
          </Card>
        </TabsContent>
      </Tabs>
    </div>
  );
};

export default CodeReview;