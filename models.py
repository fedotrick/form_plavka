from dataclasses import dataclass
from datetime import datetime, time
from typing import Optional
from constants import ExperimentType

@dataclass
class SectorData:
    """Данные сектора плавки"""
    sector_number: str
    heating_time: Optional[time] = None
    movement_time: Optional[time] = None
    pouring_time: Optional[time] = None
    temperature: Optional[float] = None

@dataclass
class MeltRecord:
    """Запись о плавке"""
    id: str
    uchet_number: str
    date: datetime
    plavka_number: str
    cluster_number: Optional[str] = None
    senior_shift: Optional[str] = None
    participant1: Optional[str] = None
    participant2: Optional[str] = None
    participant3: Optional[str] = None
    participant4: Optional[str] = None
    casting_name: Optional[str] = None
    experiment_type: Optional[ExperimentType] = None
    comment: Optional[str] = None
    created_at: Optional[datetime] = None
    
    # Данные секторов
    sector_a: Optional[SectorData] = None
    sector_b: Optional[SectorData] = None
    sector_c: Optional[SectorData] = None
    sector_d: Optional[SectorData] = None

    def to_dict(self) -> dict:
        """Преобразует запись в словарь для сохранения в БД"""
        result = {
            'id': self.id,
            'uchet_number': self.uchet_number,
            'date': self.date.strftime('%Y-%m-%d'),
            'plavka_number': self.plavka_number,
            'cluster_number': self.cluster_number,
            'senior_shift': self.senior_shift,
            'participant1': self.participant1,
            'participant2': self.participant2,
            'participant3': self.participant3,
            'participant4': self.participant4,
            'casting_name': self.casting_name,
            'experiment_type': self.experiment_type.value if self.experiment_type else None,
            'comment': self.comment
        }
        
        # Добавляем данные секторов
        for sector in ['a', 'b', 'c', 'd']:
            sector_data = getattr(self, f'sector_{sector}')
            if sector_data:
                result[f'sector_{sector}'] = sector_data.sector_number
                result[f'heating_time_{sector}'] = sector_data.heating_time.strftime('%H:%M') if sector_data.heating_time else None
                result[f'movement_time_{sector}'] = sector_data.movement_time.strftime('%H:%M') if sector_data.movement_time else None
                result[f'pouring_time_{sector}'] = sector_data.pouring_time.strftime('%H:%M') if sector_data.pouring_time else None
                result[f'temperature_{sector}'] = sector_data.temperature
                
        return result

    @classmethod
    def from_dict(cls, data: dict) -> 'MeltRecord':
        """Создает запись из словаря данных из БД"""
        # Преобразуем дату из строки
        date = datetime.strptime(data['date'], '%Y-%m-%d') if data['date'] else None
        
        # Преобразуем тип эксперимента в enum
        experiment_type = ExperimentType(data['experiment_type']) if data['experiment_type'] else None
        
        # Создаем объекты секторов
        sectors = {}
        for sector in ['a', 'b', 'c', 'd']:
            if any(data.get(f'{key}_{sector}') for key in ['sector', 'heating_time', 'movement_time', 'pouring_time', 'temperature']):
                sectors[f'sector_{sector}'] = SectorData(
                    sector_number=data.get(f'sector_{sector}'),
                    heating_time=datetime.strptime(data[f'heating_time_{sector}'], '%H:%M').time() if data.get(f'heating_time_{sector}') else None,
                    movement_time=datetime.strptime(data[f'movement_time_{sector}'], '%H:%M').time() if data.get(f'movement_time_{sector}') else None,
                    pouring_time=datetime.strptime(data[f'pouring_time_{sector}'], '%H:%M').time() if data.get(f'pouring_time_{sector}') else None,
                    temperature=float(data[f'temperature_{sector}']) if data.get(f'temperature_{sector}') else None
                )
        
        return cls(
            id=data['id'],
            uchet_number=data['uchet_number'],
            date=date,
            plavka_number=data['plavka_number'],
            cluster_number=data['cluster_number'],
            senior_shift=data['senior_shift'],
            participant1=data['participant1'],
            participant2=data['participant2'],
            participant3=data['participant3'],
            participant4=data['participant4'],
            casting_name=data['casting_name'],
            experiment_type=experiment_type,
            comment=data['comment'],
            created_at=datetime.fromisoformat(data['created_at']) if data.get('created_at') else None,
            **sectors
        )
