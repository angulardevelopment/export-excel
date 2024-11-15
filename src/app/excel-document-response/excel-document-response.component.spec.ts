import { ComponentFixture, TestBed } from '@angular/core/testing';

import { ExcelDocumentResponseComponent } from './excel-document-response.component';

describe('ExcelDocumentResponseComponent', () => {
  let component: ExcelDocumentResponseComponent;
  let fixture: ComponentFixture<ExcelDocumentResponseComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ ExcelDocumentResponseComponent ]
    })
    .compileComponents();
  });

  beforeEach(() => {
    fixture = TestBed.createComponent(ExcelDocumentResponseComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
