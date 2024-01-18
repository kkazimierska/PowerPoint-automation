from pydantic import BaseModel
class Employee(BaseModel):
    name: str
    delivery_org1: str
    delivery_org2: str
    education: str
    postion: str

Slide5 = {
    0: "empty",
    1: "Name+Initials",
    2: "Delivery Org 1",
    3: "Education",
    4: "Experience",
    5: "_professional_competencies",
    6: "Professional Interest",
    7: "Position"
}