import { ComponentFixture, TestBed } from '@angular/core/testing';

import { ExceljsComponent } from './exceljs.component';

describe('ExceljsComponent', () => {
  let component: ExceljsComponent;
  let fixture: ComponentFixture<ExceljsComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ ExceljsComponent ]
    })
    .compileComponents();
  });

  beforeEach(() => {
    fixture = TestBed.createComponent(ExceljsComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
