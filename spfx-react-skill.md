# SPFx React Development Skill

## Description
This skill provides comprehensive standards and patterns for developing SharePoint Framework (SPFx) web parts using React, PnPjs, TypeScript, and USWDS styling. Use this skill when building or refactoring SPFx solutions to ensure consistent architecture, code quality, and accessibility compliance.

## When to Use
- Building new SPFx web parts
- Refactoring existing SPFx components
- Code reviews for SPFx projects
- Teaching team members SPFx best practices
- Ensuring WCAG 2.1 Level AA compliance in SharePoint solutions
- Troubleshooting common SPFx and PnPjs issues

---

# SPFx Development Standards

You are helping develop SharePoint Framework (SPFx) web parts following these architectural standards:

## Architecture Pattern

### Service Layer
- All data fetching, SharePoint API calls, and business logic must be in separate service files
- Services should be located in a `/services` directory
- **Use PnPjs (@pnp/sp) for all SharePoint API interactions** - no direct REST calls
- Services return typed data using TypeScript interfaces
- No direct SP API calls should exist in components
- Services should handle error states and return consistent response shapes
- Initialize PnPjs properly with SPFx context

Example service structure:
```typescript
// services/ItemService.ts
import { SPFI } from '@pnp/sp';
import { getSP } from '../pnpjsConfig';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

export interface IItem {
  Id: number;
  Title: string;
}

export class ItemService {
  private _sp: SPFI;

  constructor() {
    this._sp = getSP();
  }
  
  public async getItems(listTitle: string): Promise<IItem[]> {
    try {
      const items = await this._sp.web.lists
        .getByTitle(listTitle)
        .items
        .select('Id', 'Title')
        .top(100)();
      
      return items;
    } catch (error) {
      console.error('Error fetching items:', error);
      throw error;
    }
  }

  public async getItemById(listTitle: string, itemId: number): Promise<IItem> {
    try {
      const item = await this._sp.web.lists
        .getByTitle(listTitle)
        .items
        .getById(itemId)
        .select('Id', 'Title')();
      
      return item;
    } catch (error) {
      console.error(`Error fetching item ${itemId}:`, error);
      throw error;
    }
  }

  public async createItem(listTitle: string, data: Partial<IItem>): Promise<IItem> {
    try {
      const result = await this._sp.web.lists
        .getByTitle(listTitle)
        .items
        .add(data);
      
      return result.data;
    } catch (error) {
      console.error('Error creating item:', error);
      throw error;
    }
  }

  public async updateItem(listTitle: string, itemId: number, data: Partial<IItem>): Promise<void> {
    try {
      await this._sp.web.lists
        .getByTitle(listTitle)
        .items
        .getById(itemId)
        .update(data);
    } catch (error) {
      console.error(`Error updating item ${itemId}:`, error);
      throw error;
    }
  }

  public async deleteItem(listTitle: string, itemId: number): Promise<void> {
    try {
      await this._sp.web.lists
        .getByTitle(listTitle)
        .items
        .getById(itemId)
        .delete();
    } catch (error) {
      console.error(`Error deleting item ${itemId}:`, error);
      throw error;
    }
  }
}
```

### PnPjs Configuration
- Create a centralized PnPjs configuration file in the project root
- Initialize PnPjs once with SPFx context in the web part's onInit method
- Use selective imports to reduce bundle size

Example PnPjs config:
```typescript
// pnpjsConfig.ts
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFI, SPFx } from '@pnp/sp';

let _sp: SPFI;

export const getSP = (): SPFI => {
  return _sp;
};

export const setSP = (context: WebPartContext): void => {
  _sp = spfi().using(SPFx(context));
};
```

Example web part initialization:
```typescript
// MyWebPart.ts
import { setSP } from '../../pnpjsConfig';

export default class MyWebPart extends BaseClientSideWebPart<IMyWebPartProps> {
  protected async onInit(): Promise<void> {
    await super.onInit();
    
    // Initialize PnPjs
    setSP(this.context);
  }

  public render(): void {
    const element: React.ReactElement<IMyComponentProps> = React.createElement(
      MyComponent,
      {
        // Pass services, not context
      }
    );

    ReactDom.render(element, this.domElement);
  }
}
```

### React Components
- Use **functional components only** - no class components
- Use React hooks (useState, useEffect, useCallback, useMemo, etc.)
- Keep components focused and single-responsibility
- Extract reusable logic into custom hooks when appropriate
- Props should be typed with TypeScript interfaces
- Components receive service instances via props, not context

### Styling
- Use **USWDS (U.S. Web Design System) classes exclusively** for styling
- No inline styles or custom CSS unless absolutely necessary
- Follow USWDS component patterns and accessibility guidelines
- Leverage USWDS utility classes for spacing, typography, and layout

## ESLint & Code Quality Standards

Follow standard SPFx ESLint rules including:

### TypeScript/JavaScript Rules
- **No `any` types** - always use specific types or proper generics
- **Explicit return types** on functions (especially public methods and exported functions)
- **No unused variables or imports** - clean up all unused code
- **Prefer const over let** - use `const` by default, `let` only when reassignment is needed
- **No var declarations** - always use `const` or `let`
- **Semicolons required** - end statements with semicolons
- **Single quotes for strings** - use `'string'` not `"string"`
- **No console statements in production** - use proper logging or remove before committing (exception: error logging in services)

### React-Specific Rules
- **Hook dependencies must be complete** - `useEffect`, `useCallback`, `useMemo` must list all dependencies
- **No missing keys in lists** - always provide unique `key` prop when mapping arrays
- **Hooks only at top level** - no conditional hook calls
- **Props destructuring preferred** - destructure props in function signature when practical
- **PascalCase for components** - component names must be PascalCase
- **Interfaces prefixed with I** - `IMyComponentProps`, `IItem`, etc.

### PnPjs Best Practices
- **Selective imports** - only import the specific PnP modules you need (reduces bundle size)
- **Use select() to limit returned fields** - don't fetch all columns if you only need a few
- **Use top() for pagination** - limit results appropriately
- **Proper error handling** - wrap PnP calls in try/catch blocks
- **Batching for multiple operations** - use PnPjs batching for multiple create/update operations
- **Initialize once** - PnPjs should be initialized once in onInit, not in services

### Accessibility Rules
- **Images must have alt text** - `<img>` tags require `alt` attribute
- **Interactive elements must be keyboard accessible** - proper use of semantic HTML
- **ARIA attributes used correctly** - follow ARIA best practices
- **Sufficient color contrast** - meet WCAG AA standards (USWDS helps with this)

### File Organization
- **One component per file** - don't export multiple components from same file
- **Index files for cleaner imports** - use index.ts for barrel exports when appropriate
- **Consistent naming** - file names should match component/class names

Example component following all rules:
```typescript
import * as React from 'react';
import { ItemService, IItem } from '../services/ItemService';

export interface IMyComponentProps {
  itemService: ItemService;
  listTitle: string;
}

export const MyComponent: React.FC<IMyComponentProps> = ({ itemService, listTitle }) => {
  const [items, setItems] = React.useState<IItem[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string | null>(null);

  React.useEffect(() => {
    const fetchData = async (): Promise<void> => {
      try {
        setLoading(true);
        const data = await itemService.getItems(listTitle);
        setItems(data);
        setError(null);
      } catch (err) {
        setError('Failed to load items');
      } finally {
        setLoading(false);
      }
    };
    
    fetchData();
  }, [itemService, listTitle]);

  const handleDelete = React.useCallback(async (itemId: number): Promise<void> => {
    try {
      await itemService.deleteItem(listTitle, itemId);
      setItems(items.filter((item) => item.Id !== itemId));
    } catch (err) {
      setError('Failed to delete item');
    }
  }, [itemService, listTitle, items]);

  if (loading) {
    return <div className="usa-loader" aria-live="polite">Loading...</div>;
  }

  if (error) {
    return (
      <div className="usa-alert usa-alert--error" role="alert">
        <div className="usa-alert__body">
          <p className="usa-alert__text">{error}</p>
        </div>
      </div>
    );
  }

  return (
    <div className="grid-container">
      <h2 className="font-heading-xl margin-bottom-2">Items from {listTitle}</h2>
      <ul className="usa-list">
        {items.map((item) => (
          <li key={item.Id} className="margin-bottom-1">
            <span>{item.Title}</span>
            <button
              className="usa-button usa-button--secondary margin-left-1"
              onClick={() => handleDelete(item.Id)}
              type="button"
            >
              Delete
            </button>
          </li>
        ))}
      </ul>
    </div>
  );
};
```

## Key Principles
- Separation of concerns: UI components should not know about SharePoint APIs or PnPjs
- Services encapsulate all PnPjs interactions
- PnPjs is initialized once in web part onInit
- Services are injected into components via props (dependency injection pattern)
- All async operations use async/await syntax with proper error handling
- Error boundaries and error states should be handled gracefully
- Accessibility is critical - leverage USWDS's built-in WCAG compliance
- Code must pass ESLint with zero warnings before committing
- Type safety is non-negotiable - leverage TypeScript fully
- Bundle size matters - use selective PnP imports

## Common PnPjs Patterns

### Querying with Filters
```typescript
const items = await this._sp.web.lists
  .getByTitle(listTitle)
  .items
  .filter(`Status eq 'Active'`)
  .select('Id', 'Title', 'Status')
  .top(50)();
```

### Expanding Lookup Fields
```typescript
const items = await this._sp.web.lists
  .getByTitle(listTitle)
  .items
  .select('Id', 'Title', 'Author/Title', 'Author/EMail')
  .expand('Author')();
```

### Batching Operations
```typescript
const [batchedSP, execute] = this._sp.batched();

items.forEach((item) => {
  batchedSP.web.lists.getByTitle(listTitle).items.add(item);
});

await execute();
```

## Troubleshooting Patterns

### Problem: "Cannot read property of undefined" when calling PnPjs
**Symptoms**: Error occurs when trying to use `this._sp` in a service
**Cause**: PnPjs wasn't initialized before the service tried to use it
**Solution**:
```typescript
// In your web part's onInit, ensure setSP is called
protected async onInit(): Promise<void> {
  await super.onInit();
  setSP(this.context); // This must happen before any service calls
}

// In your service, add a safety check
export class ItemService {
  private _sp: SPFI;

  constructor() {
    this._sp = getSP();
    if (!this._sp) {
      throw new Error('PnPjs has not been initialized. Call setSP() in your web part onInit.');
    }
  }
}
```

### Problem: "Module not found" errors for @pnp packages
**Symptoms**: Build fails with errors like `Cannot find module '@pnp/sp/webs'`
**Cause**: Missing selective imports or PnP packages not installed
**Solution**:
```bash
# Install the required PnP packages
npm install @pnp/sp @pnp/core --save

# Ensure you have the selective imports at the top of your service
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
```

### Problem: "React Hook useEffect has a missing dependency" ESLint warning
**Symptoms**: ESLint warns about missing dependencies in useEffect
**Cause**: Dependencies array doesn't include all values used in the effect
**Solution**:
```typescript
// BAD - missing itemService dependency
React.useEffect(() => {
  const fetchData = async (): Promise<void> => {
    const data = await itemService.getItems(listTitle);
    setItems(data);
  };
  fetchData();
}, [listTitle]); // ESLint will warn about missing itemService

// GOOD - all dependencies included
React.useEffect(() => {
  const fetchData = async (): Promise<void> => {
    const data = await itemService.getItems(listTitle);
    setItems(data);
  };
  fetchData();
}, [itemService, listTitle]); // All dependencies present
```

### Problem: Infinite re-render loop with useEffect
**Symptoms**: Component re-renders continuously, browser becomes unresponsive
**Cause**: Dependency in useEffect changes on every render (like creating new objects/functions)
**Solution**:
```typescript
// BAD - creates new service instance every render
const MyComponent: React.FC = () => {
  const service = new ItemService(); // New instance every render
  
  React.useEffect(() => {
    service.getItems(); // Triggers re-render, creates new service, infinite loop
  }, [service]); // service is always "new"
}

// GOOD - service passed as prop or memoized
const MyComponent: React.FC<IMyComponentProps> = ({ itemService }) => {
  React.useEffect(() => {
    itemService.getItems();
  }, [itemService]); // itemService reference is stable
}

// ALTERNATIVE - use useCallback for functions
const handleFetch = React.useCallback(async () => {
  await itemService.getItems(listTitle);
}, [itemService, listTitle]);

React.useEffect(() => {
  handleFetch();
}, [handleFetch]); // handleFetch reference is stable
```

### Problem: "401 Unauthorized" or "403 Forbidden" from PnPjs calls
**Symptoms**: SharePoint API calls fail with permission errors
**Cause**: Insufficient permissions or incorrect context
**Solution**:
```typescript
// 1. Verify your web part has the correct permissions in package-solution.json
{
  "webApiPermissionRequests": [
    {
      "resource": "SharePoint",
      "scope": "Web.Read"
    }
  ]
}

// 2. Check if you're using the correct context
// Make sure you're passing WebPartContext, not a different context type
export const setSP = (context: WebPartContext): void => {
  _sp = spfi().using(SPFx(context));
};

// 3. For cross-site queries, ensure proper permissions
const items = await this._sp.web
  .getList('/sites/othersite/lists/mylist') // May require additional permissions
  .items();
```

### Problem: Build succeeds but web part doesn't appear in workbench
**Symptoms**: `gulp serve` succeeds, but web part is missing from the toolbox
**Cause**: Manifest issues or serving from wrong URL
**Solution**:
```typescript
// 1. Check your web part manifest (MyWebPart.manifest.json)
// Ensure preconfiguredEntries exists and is valid
{
  "preconfiguredEntries": [{
    "groupId": "5c03119e-3074-46fd-976b-c60198311f70",
    "group": { "default": "Other" },
    "title": { "default": "My Web Part" },
    "description": { "default": "My web part description" }
  }]
}

// 2. Clear browser cache and restart gulp serve
// 3. Make sure you're accessing the correct workbench URL
// https://yourtenant.sharepoint.com/_layouts/15/workbench.aspx
```

### Problem: TypeError with PnPjs when updating items
**Symptoms**: `TypeError: Cannot read property 'getById' of undefined`
**Cause**: Chaining methods incorrectly or missing imports
**Solution**:
```typescript
// BAD - missing items() call before getById
await this._sp.web.lists
  .getByTitle(listTitle)
  .getById(itemId) // ERROR - getById doesn't exist on list
  .update(data);

// GOOD - proper chaining with items
await this._sp.web.lists
  .getByTitle(listTitle)
  .items // Need .items before .getById
  .getById(itemId)
  .update(data);

// Also ensure you have the import
import '@pnp/sp/items';
```

### Problem: Stale data after create/update operations
**Symptoms**: UI doesn't reflect newly created/updated items
**Cause**: Component state not updated after successful operation
**Solution**:
```typescript
// BAD - state not updated after creation
const handleCreate = async (title: string): Promise<void> => {
  await itemService.createItem(listTitle, { Title: title });
  // UI still shows old data
};

// GOOD - refresh data after operation
const handleCreate = async (title: string): Promise<void> => {
  const newItem = await itemService.createItem(listTitle, { Title: title });
  setItems([...items, newItem]); // Update local state
};

// ALTERNATIVE - refetch all data
const handleCreate = async (title: string): Promise<void> => {
  await itemService.createItem(listTitle, { Title: title });
  const refreshedItems = await itemService.getItems(listTitle);
  setItems(refreshedItems);
};
```

### Problem: "Property 'X' does not exist on type" TypeScript errors
**Symptoms**: TypeScript complains about properties that exist in SharePoint
**Cause**: Interface doesn't match SharePoint column schema
**Solution**:
```typescript
// BAD - interface missing custom columns
export interface IItem {
  Id: number;
  Title: string;
  // Missing custom columns
}

// GOOD - interface matches all columns being queried
export interface IItem {
  Id: number;
  Title: string;
  CustomField: string;
  Department: string;
  CreatedDate: string;
}

// When selecting fields, ensure they match your interface
const items = await this._sp.web.lists
  .getByTitle(listTitle)
  .items
  .select('Id', 'Title', 'CustomField', 'Department', 'CreatedDate')();
```

### Problem: Web part crashes in SharePoint but works in workbench
**Symptoms**: Web part works locally but fails when deployed to SharePoint
**Cause**: Often related to CORS, context differences, or missing dependencies
**Solution**:
```typescript
// 1. Check browser console for CORS errors
// 2. Verify all dependencies are bundled correctly (check package.json)
// 3. Ensure external scripts are loaded from approved CDNs in config.json

// 4. Add defensive coding for context differences
protected onInit(): Promise<void> {
  return super.onInit().then(() => {
    try {
      setSP(this.context);
    } catch (error) {
      console.error('Failed to initialize PnPjs:', error);
      // Handle gracefully
    }
  });
}

// 5. Test with --ship flag to simulate production bundle
gulp bundle --ship
gulp package-solution --ship
```

### Problem: Performance issues with large lists
**Symptoms**: Slow load times or timeouts when fetching data
**Cause**: Fetching too many items or not using pagination
**Solution**:
```typescript
// BAD - fetching all items at once
const items = await this._sp.web.lists
  .getByTitle(listTitle)
  .items(); // Could return thousands of items

// GOOD - use top() and pagination
const items = await this._sp.web.lists
  .getByTitle(listTitle)
  .items
  .top(100)(); // Limit to 100 items

// BETTER - implement proper pagination
public async getItemsPaged(listTitle: string, pageSize: number = 100): Promise<IPagedItems> {
  const result = await this._sp.web.lists
    .getByTitle(listTitle)
    .items
    .select('Id', 'Title')
    .top(pageSize)
    .getPaged();
    
  return {
    items: result.results,
    hasNext: result.hasNext,
    getNext: result.hasNext ? () => result.getNext() : null
  };
}

// BEST - use select() to only fetch needed fields
const items = await this._sp.web.lists
  .getByTitle(listTitle)
  .items
  .select('Id', 'Title') // Only fetch what you need
  .top(100)();
```

### Problem: ESLint errors prevent build
**Symptoms**: Build fails with ESLint violations
**Solution**:
```bash
# View all ESLint errors
npm run lint

# Auto-fix fixable issues
npm run lint -- --fix

# Common fixes:
# 1. Remove unused imports
# 2. Add explicit return types
# 3. Fix dependency arrays in hooks
# 4. Remove console.log statements
# 5. Add semicolons
# 6. Change double quotes to single quotes
```

## Testing Patterns

### Unit Testing Services
```typescript
// ItemService.test.ts
import { ItemService } from './ItemService';
import { SPFI } from '@pnp/sp';

describe('ItemService', () => {
  let service: ItemService;
  let mockSP: jest.Mocked<SPFI>;

  beforeEach(() => {
    // Mock PnPjs
    mockSP = {
      web: {
        lists: {
          getByTitle: jest.fn().mockReturnValue({
            items: {
              select: jest.fn().mockReturnThis(),
              top: jest.fn().mockReturnValue(Promise.resolve([]))
            }
          })
        }
      }
    } as any;

    service = new ItemService();
    (service as any)._sp = mockSP;
  });

  it('should fetch items successfully', async () => {
    const mockItems = [{ Id: 1, Title: 'Test' }];
    mockSP.web.lists.getByTitle('TestList').items.top = jest.fn().mockReturnValue(Promise.resolve(mockItems));

    const result = await service.getItems('TestList');
    
    expect(result).toEqual(mockItems);
    expect(mockSP.web.lists.getByTitle).toHaveBeenCalledWith('TestList');
  });
});
```

### Component Testing
```typescript
// MyComponent.test.tsx
import * as React from 'react';
import { render, screen, waitFor } from '@testing-library/react';
import { MyComponent } from './MyComponent';
import { ItemService } from '../services/ItemService';

describe('MyComponent', () => {
  let mockService: jest.Mocked<ItemService>;

  beforeEach(() => {
    mockService = {
      getItems: jest.fn().mockResolvedValue([
        { Id: 1, Title: 'Item 1' },
        { Id: 2, Title: 'Item 2' }
      ])
    } as any;
  });

  it('renders items after loading', async () => {
    render(<MyComponent itemService={mockService} listTitle="TestList" />);

    expect(screen.getByText('Loading...')).toBeInTheDocument();

    await waitFor(() => {
      expect(screen.getByText('Item 1')).toBeInTheDocument();
      expect(screen.getByText('Item 2')).toBeInTheDocument();
    });
  });

  it('displays error message on failure', async () => {
    mockService.getItems.mockRejectedValue(new Error('API Error'));

    render(<MyComponent itemService={mockService} listTitle="TestList" />);

    await waitFor(() => {
      expect(screen.getByText('Failed to load items')).toBeInTheDocument();
    });
  });
});
```

When generating code, always follow these patterns unless explicitly instructed otherwise.
