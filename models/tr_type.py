from sqlalchemy import Column, Integer, String, Boolean, ForeignKey, JSON, Text, UniqueConstraint
from sqlalchemy.orm import declarative_base, relationship

Base = declarative_base()


class TrType(Base):
    __tablename__ = "tr_types"
    id = Column(Integer, primary_key=True)
    name = Column(String(100), nullable=False, unique=True)


class VoltageClass(Base):
    __tablename__ = "voltage_classes"
    id = Column(Integer, primary_key=True)
    value = Column(String(50), nullable=False, unique=True)


class Climat(Base):
    __tablename__ = "climats"
    id = Column(Integer, primary_key=True)
    name = Column(String(50), nullable=False, unique=True)


class IsolType(Base):
    __tablename__ = "isol_types"
    id = Column(Integer, primary_key=True)
    name = Column(String(100), nullable=False, unique=True)


class IsolColor(Base):
    __tablename__ = "isol_colors"
    id = Column(Integer, primary_key=True)
    name = Column(String(50), nullable=False, unique=True)


class AccuracyClass(Base):
    __tablename__ = "accuracy_classes"
    id = Column(Integer, primary_key=True)
    name = Column(String(20), nullable=False, unique=True)
    is_measuring = Column(Boolean, default=False)


class TrTypeRule(Base):
    __tablename__ = "tr_type_rules"
    id = Column(Integer, primary_key=True, autoincrement=True)
    tr_type_id = Column(Integer, ForeignKey("tr_types.id"), nullable=False, unique=True)
    voltage_classes = Column(JSON, default=list)
    climats = Column(JSON, default=list)
    isol_types = Column(JSON, default=list)
    isol_colors = Column(JSON, default=list)
    accuracy_classes = Column(JSON, default=list)
    tr_type = relationship("TrType")


class KnowledgeEntry(Base):
    """Расширяемая БЗ: любой параметр можно добавить без изменения схемы Python."""
    __tablename__ = "knowledge_entries"
    id = Column(Integer, primary_key=True, autoincrement=True)
    category = Column(String(80), nullable=False, default="product")
    key = Column(String(150), nullable=False)
    value = Column(Text, nullable=False)
    aliases = Column(JSON, default=list)
    tr_type_id = Column(Integer, ForeignKey("tr_types.id"), nullable=True)
    voltage = Column(String(50), nullable=True)
    source = Column(String(255), nullable=False, default="manual")
    notes = Column(Text, nullable=True)
    active = Column(Boolean, nullable=False, default=True)
    tr_type = relationship("TrType")

    __table_args__ = (
        UniqueConstraint("category", "key", "tr_type_id", "voltage", name="uq_knowledge_scope"),
    )


class FieldRule(Base):
    """Связывает реальное название строки ТЗ с каноническим ключом БЗ."""
    __tablename__ = "field_rules"
    id = Column(Integer, primary_key=True, autoincrement=True)
    canonical_name = Column(String(150), nullable=False, unique=True)
    aliases = Column(JSON, default=list)
    data_type = Column(String(30), nullable=False, default="string")
    db_key = Column(String(150), nullable=True)
    section = Column(String(80), nullable=True)
    priority = Column(Integer, nullable=False, default=100)
    active = Column(Boolean, nullable=False, default=True)


class Tender(Base):
    """Сохраненный шаблон тендерного документа."""
    __tablename__ = "tenders"
    id = Column(Integer, primary_key=True, autoincrement=True)
    filename = Column(String(255), nullable=True)
    object_name = Column(String(255), nullable=True)
    quantity = Column(String(100), nullable=True)
    delivery_date = Column(String(100), nullable=True)
    delivery_address = Column(String(500), nullable=True)
    created_at = Column(String(50), nullable=True)
    specialist_name = Column(String(255), nullable=True)


class TenderParameter(Base):
    """Строка сохраненного шаблона с ответом участника."""
    __tablename__ = "tender_parameters"
    id = Column(Integer, primary_key=True, autoincrement=True)
    tender_id = Column(Integer, ForeignKey("tenders.id"), nullable=False)
    section_number = Column(String(50), nullable=True)
    parameter_name = Column(String(500), nullable=False)
    required_value = Column(Text, nullable=True)
    proposed_value = Column(Text, nullable=True)
    row_index = Column(Integer, nullable=True)
    table_index = Column(Integer, nullable=True)
    target_cell_index = Column(Integer, nullable=True)
    field_key = Column(String(150), nullable=True)
    is_mandatory = Column(Boolean, default=False)
    tender = relationship("Tender")


class ParameterProfile(Base):
    """Профиль типового изделия/напряжения из старой БЗ."""
    __tablename__ = "parameter_profiles"
    id = Column(Integer, primary_key=True, autoincrement=True)
    name = Column(String(255), nullable=False)
    voltage = Column(String(50), nullable=True)
    description = Column(Text, nullable=True)


class ProfileParameter(Base):
    __tablename__ = "profile_parameters"
    id = Column(Integer, primary_key=True, autoincrement=True)
    profile_id = Column(Integer, ForeignKey("parameter_profiles.id"), nullable=False)
    parameter_name = Column(String(500), nullable=False)
    value = Column(Text, nullable=True)
    profile = relationship("ParameterProfile")
