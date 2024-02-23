from datetime import datetime

from sqlalchemy import create_engine, Column, Integer, DATETIME, VARCHAR, TIME, Boolean, and_, DATE
from sqlalchemy.orm import sessionmaker, declarative_base

engine = create_engine('sqlite:///mydb.db', connect_args={"check_same_thread": False})
session = sessionmaker(bind=engine)()

Base = declarative_base()


class BaseModel(Base):
    __abstract__ = True

    id = Column(Integer, nullable=False, unique=True, primary_key=True, autoincrement=True)

    def __repr__(self):
        return f"<{__class__.__name__}(id={self.id})>"


class Task(BaseModel):
    __tablename__ = 'tasks'

    day = Column(Integer)
    month = Column(Integer)
    year = Column(Integer)
    name = Column(VARCHAR(255))
    time = Column(TIME)
    duration = Column(Integer)
    price = Column(Integer)
    is_finished = Column(Boolean)

    def __init__(self, date: datetime.date, name, time, duration, price, is_finished=False):
        self.day, self.month, self.year = date.day, date.month, date.year
        self.name = name
        self.time = time
        self.duration = duration
        self.price = price
        self.is_finished = is_finished

    @staticmethod
    def add_task(task: list):
        try:
            new_task = Task(*task)
            if session.query(Task).filter(
                    and_(Task.year == task[0].year,
                         Task.month == task[0].month,
                         Task.day == task[0].day,
                         Task.time == task[2])).first():
                return False
            session.add(new_task)
            session.commit()
            return True
        except Exception as ex:
            print('Error in func add_task -', ex)
            session.rollback()
            return False

    @staticmethod
    def get_unfinished_tasks():
        return session.query(Task).filter(Task.is_finished == False).order_by(Task.year, Task.month, Task.day,
                                                                              Task.time).all()

    @staticmethod
    def delete_task(pk):
        try:
            task = session.get(Task, pk)
            if task:
                session.delete(task)
                session.commit()
        except Exception as ex:
            print('Error in func delete_task -', ex)
            session.rollback()

    @staticmethod
    def finish_task(pk):
        try:
            task = session.get(Task, pk)
            if task:
                task.is_finished = True
                session.commit()
        except Exception as ex:
            print('Error in func finish_task -', ex)
            session.rollback()

    @staticmethod
    def get_by_day(date: datetime.date):
        return session.query(Task).filter(
            and_(Task.year == date.year, Task.month == date.month, Task.day == date.day)).order_by(Task.time).all()

    @staticmethod
    def get_by_month(date: datetime.date):
        return session.query(Task).filter(
            and_(Task.year == date.year, Task.month == date.month)).order_by(Task.day, Task.time).all()

    @staticmethod
    def get_by_year(date: datetime.date):
        return session.query(Task).filter(Task.year == date.year).order_by(Task.month, Task.day, Task.time).all()

    @staticmethod
    def get_all_tasks():
        return session.query(Task).order_by(Task.year, Task.month, Task.day, Task.time).all()


if __name__ == '__main__':
    Base.metadata.create_all(engine)
